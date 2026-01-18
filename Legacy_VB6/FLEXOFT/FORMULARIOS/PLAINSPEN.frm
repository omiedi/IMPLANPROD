VERSION 5.00
Begin VB.Form PLAINSPEN 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hoja de Especificacion de Inspección"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   11910
   Begin VB.TextBox Text7 
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
      Left            =   1680
      TabIndex        =   25
      Top             =   1785
      Width           =   2115
   End
   Begin VB.TextBox Text8 
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
      Left            =   7140
      TabIndex        =   24
      Top             =   1800
      Width           =   2115
   End
   Begin VB.TextBox Text9 
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
      Left            =   4515
      TabIndex        =   23
      Top             =   1785
      Width           =   705
   End
   Begin VB.TextBox Text10 
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
      Left            =   9975
      TabIndex        =   22
      Top             =   1785
      Width           =   705
   End
   Begin VB.CommandButton Command7 
      Caption         =   " "
      Height          =   260
      Left            =   1680
      TabIndex        =   18
      Top             =   8190
      Width           =   4100
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
      Height          =   855
      Left            =   10920
      Picture         =   "PLAINSPEN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2650
      Width           =   855
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Selecc Todo"
      Height          =   260
      Left            =   9435
      TabIndex        =   14
      Top             =   8190
      Width           =   1290
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Pegar"
      Height          =   260
      Left            =   8145
      TabIndex        =   13
      Top             =   8190
      Width           =   1275
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Copiar"
      Height          =   260
      Left            =   6975
      TabIndex        =   12
      Top             =   8190
      Width           =   1170
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cortar"
      Height          =   260
      Left            =   5790
      TabIndex        =   11
      Top             =   8190
      Width           =   1170
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5580
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   2625
      Width           =   9060
   End
   Begin VB.CommandButton Command5 
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
      Height          =   800
      Left            =   10920
      Picture         =   "PLAINSPEN.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprime Ficha de Ingenieria"
      Top             =   6930
      Width           =   855
   End
   Begin VB.CommandButton Command2 
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
      Left            =   10920
      Picture         =   "PLAINSPEN.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1785
      Width           =   855
   End
   Begin VB.CommandButton Command9 
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
      Left            =   10920
      Picture         =   "PLAINSPEN.frx":0DB6
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Sale de la Aplicación"
      Top             =   105
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K."
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
      Left            =   10920
      Picture         =   "PLAINSPEN.frx":0F00
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Graba las Modificaciones"
      Top             =   945
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   10740
      X2              =   10740
      Y1              =   2415
      Y2              =   8430
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10340
      Picture         =   "PLAINSPEN.frx":120A
      Top             =   7930
      Width           =   2010
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   6510
      TabIndex        =   29
      Top             =   420
      Width           =   2430
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rev:"
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
      Left            =   3780
      TabIndex        =   28
      Top             =   1850
      Width           =   645
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rev:"
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
      Left            =   9240
      TabIndex        =   27
      Top             =   1850
      Width           =   645
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "QC-HRI Nro:"
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
      Left            =   105
      TabIndex        =   26
      Top             =   1850
      Width           =   1485
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   9975
      TabIndex        =   21
      Top             =   420
      Width           =   750
   End
   Begin VB.Label CONTROCAM 
      BackStyle       =   0  'Transparent
      Caption         =   "Label17"
      Height          =   330
      Left            =   105
      TabIndex        =   20
      Top             =   105
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rev:"
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
      Left            =   9240
      TabIndex        =   19
      Top             =   525
      Width           =   645
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00800000&
      Height          =   555
      Left            =   1680
      TabIndex        =   1
      Top             =   1050
      Width           =   9045
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   1680
      TabIndex        =   0
      Top             =   420
      Width           =   2430
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Item  Descripcion de la Operacion              Dimension  Instrum   Frec"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   1680
      TabIndex        =   17
      Top             =   2380
      Width           =   9045
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "QA-PIE Nro:"
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
      Left            =   5565
      TabIndex        =   16
      Top             =   1850
      Width           =   1485
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Plano.Nro:"
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
      Left            =   5145
      TabIndex        =   15
      Top             =   525
      Width           =   1275
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
      TabIndex        =   9
      Top             =   2625
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
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
      Left            =   315
      TabIndex        =   8
      Top             =   525
      Width           =   1275
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion Corta:"
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
      Left            =   105
      TabIndex        =   7
      Top             =   1050
      Width           =   1485
   End
End
Attribute VB_Name = "PLAINSPEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CI%, CABAS, CODE$, CIE%
Dim HIDOCLICK As Integer
Dim TOPEINI(3), IZQUINI(3), ANCHINI(3), ALTUINI(3)
Dim IMACTI
Dim MODIFK%
'
Private Sub Command1_Click()
    Call GRADESINSP
    Call GRADATINSP
    MODIFK% = 0
End Sub

Private Sub Command11_Click()
   COMIE = Text3.SelStart
   LOSEL = Text3.SelLength
   Text3.TEXT = Left$(Text3.TEXT, COMIE) + Clipboard.GetText + Mid$(Text3.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command12_Click()
   Text3.SelStart = 0
   Text3.SelLength = Len(Text3.TEXT)
   Text3.Refresh
End Sub

Private Sub Command26_Click()
   Command26.Enabled = False
   OPGRAFIC.Show 1
   Command26.Enabled = True
End Sub


Private Sub Command4_Click()
   COMIE = Text3.SelStart
   LOSEL = Text3.SelLength
   Text3.TEXT = Left$(Text3.TEXT, COMIE) + Mid$(Text3.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command5_Click()
   PRIFORM$ = TRIM$(CONTROL$("FORGRAFI", "DOCUGRA4"))
   If PRIFORM$ <> "" Then
     If EXISTE%(LUDAT$ + PRIFORM$ + ".FRM") Then
       CIXI% = CODINT%(Label7.Caption)
       Call PRIHOJRINSP(CIXI%)
       Exit Sub
     End If
   End If
   '
   ORIENTANT = Printer.Orientation
   Printer.Orientation = 2
   PRINTFORM
   Printer.Orientation = ORIENTANT
End Sub
'


Private Sub Command8_Click()
   Clipboard.SetText Mid$(Text3.TEXT, Text3.SelStart + 1, Text3.SelLength)
End Sub

Private Sub CONTROCAM_Change()
   Label7.Caption = FICHINGEPRO.Text1.TEXT
   Label9.Caption = FICHINGEPRO.Text2.TEXT
   Label8.Caption = TRIM$(FICHINGEPRO.Text7.TEXT)
   Label11.Caption = TRIM$(FICHINGEPRO.Text9.TEXT)
   Call CARDATINSP
   Call CARDESINSP
   MODIFK% = 0
End Sub


Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Screen.MousePointer = 1
    '
End Sub


Private Sub Form_Unload(Cancel As Integer)
    INI& = 1 + ULTREG&("ESTRUDOS")
    FIN& = ULTREG&("ESTRUNUE")
    For U& = INI& To FIN&
        RB$ = String$(256, 0)
        Mid$(RB$, 251, 4) = MKS$(0.0001)
        Call GRAREG("ESTRUDOS", RB$, U&)
    Next U&
    '
    Cancel = 0
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command2_Click()
    Call HELPONLINE("FICHACAL")
End Sub
'

Private Sub Command9_Click()
    If MODIFK% > 0 Then
        RESPU = MsgBox("Abandona sin Guardar los Cambios ?", 4371, "Operacion Critica")
        If RESPU <> 6 Then
            Command1.SetFocus
            Exit Sub
        End If
        MODIFK% = 0
    End If
    '
    Hide
End Sub

Sub CARDESINSP()
   '
   COD$ = Label7.Caption
   CI% = CODINT%(Label7.Caption)
   Text3.TEXT = ""
   NOPE% = Val(Label8.Caption)
   '
   If CI% > 0 Then
      If CI% <= NUMAS% Then
         Screen.MousePointer = 11
         RNC$ = RECFILT$("DESOPINS", 1, MKI$(CI%))
         URN& = ULTREG&("DESOPINS")
         '
         For U& = 1 To Len(RNC$) Step 4
            RENOTA& = CVS(Mid$(RNC$, U&, 4))
            If RENOTA& > 0 Then
               If RENOTA& <= URN& Then
                  TTXX$ = REGLEIDO$("DESOPINS", RENOTA&)
                  TTYY$ = Mid$(TTXX$, 5, 72)
                  If Text3.TEXT <> "" Then
                     Text3.TEXT = Text3.TEXT + Chr$(13) + Chr$(10)
                  End If
                  Text3.TEXT = Text3.TEXT + RTrim$(Left$(TTYY$, 72))
               End If
            End If
         Next U&
         Screen.MousePointer = 1
      End If
   End If
   Text3.Refresh
   '
End Sub

Sub GRADESINSP()
   '
   COD$ = Label7.Caption
   CI% = CODINT%(COD$)
   If CI% > 0 Then
      If CI% <= NUMAS% Then
         Screen.MousePointer = 11
         URN& = ULTREG&("DESOPINS")
         RNC$ = RECFILT$("DESOPINS", 1, MKI$(CI%))
         KK& = (-3)
         '
         TEXTO$ = Text3.TEXT + Chr$(13) + Chr$(10)
         Call FRATEXTO(TEXTO$, 72)
         For LL% = 1 To CARETEX%
            TTYY$ = MKI$(CI%) + MKI$(NOPE%) + RETEX$(LL%)
15          KK& = KK& + 4
            If KK& <= Len(RNC$) - 3 Then
                RENOTA& = CVS(Mid$(RNC$, KK&, 4))
              Else
                URN& = URN& + 1
                RENOTA& = URN&
            End If
            TTXX$ = REGLEIDO$("DESOPINS", RENOTA&)
            Call GRAREG("DESOPINS", TTYY$, RENOTA&)
         Next LL%
         '
20       While KK& < Len(RNC$) - 3
            KK& = KK& + 4
            RENOTA& = CVS(Mid$(RNC$, KK&, 4))
            Call GRAREG("DESOPINS", MKI$(0) + MKI$(0) + Space$(68), RENOTA&)
         Wend
         '
         Call CIERRARCH("DESOPINS")
         Screen.MousePointer = 1
      End If
   End If

End Sub
'

Sub CARDATINSP()
   '
   COD$ = Label7.Caption
   CI% = CODINT%(Label7.Caption)
   '
   If CI% > 0 Then
      If CI% <= NUMAS% Then
         Screen.MousePointer = 11
         RNC$ = RECFILT$("DATADINS", 1, MKI$(CI%))
         URN& = ULTREG&("DATADINS")
         '
         For U& = 1 To Len(RNC$) Step 4
            RENOTA& = CVS(Mid$(RNC$, U&, 4))
            If RENOTA& > 0 Then
               If RENOTA& <= URN& Then
                  TTXX$ = REGLEIDO$("DATADINS", RENOTA&)
                  GoTo 29
               End If
            End If
         Next U&
         TTXX$ = REGBLAN$("DATADINS")
         '
29       Text7.TEXT = TRIM$(Mid$(TTXX$, 5, 16))
         Text9.TEXT = TRIM$(Mid$(TTXX$, 21, 4))
         Text8.TEXT = TRIM$(Mid$(TTXX$, 25, 16))
         Text10.TEXT = TRIM$(Mid$(TTXX$, 41, 4))
         Screen.MousePointer = 1
      End If
   End If
   '
End Sub

Sub GRADATINSP()
   '
   COD$ = Label7.Caption
   CI% = CODINT%(COD$)
   If CI% > 0 Then
      If CI% <= NUMAS% Then
         Screen.MousePointer = 11
         URN& = ULTREG&("DATADINS")
         RNC$ = RECFILT$("DATADINS", 1, MKI$(CI%))
         KK& = (-3)
         '
         TTYY$ = REGBLAN$("DATADINS")
         Call REPLA(TTYY$, MKI$(CI%), 1, 2)
         Call REPLA(TTYY$, MKI$(0), 3, 2)
         Call REPLA(TTYY$, Text7.TEXT, 5, 16)
         Call REPLA(TTYY$, Text9.TEXT, 21, 4)
         Call REPLA(TTYY$, Text8.TEXT, 25, 16)
         Call REPLA(TTYY$, Text10.TEXT, 41, 4)
         '
15       KK& = KK& + 4
         If KK& <= Len(RNC$) - 3 Then
              RENOTA& = CVS(Mid$(RNC$, KK&, 4))
            Else
              URN& = URN& + 1
              RENOTA& = URN&
         End If
         TTXX$ = REGLEIDO$("DATADINS", RENOTA&)
         Call GRAREG("DATADINS", TTYY$, RENOTA&)
         '
20       While KK& < Len(RNC$) - 3
            KK& = KK& + 4
            RENOTA& = CVS(Mid$(RNC$, KK&, 4))
            Call GRAREG("DATADINS", MKI$(CI%) + MKI$(0) + Space$(124), RENOTA&)
         Wend
         '
         Call CIERRARCH("DATADINS")
         Screen.MousePointer = 1
      End If
   End If
   '
End Sub

Private Sub Text10_Change()
   MODIFK% = 1
End Sub

Private Sub Text3_Change()
   MODIFK% = 1
End Sub

Private Sub Text7_Change()
   MODIFK% = 1
End Sub

Private Sub Text8_Change()
   MODIFK% = 1
End Sub

Private Sub Text9_Change()
   MODIFK% = 1
End Sub
