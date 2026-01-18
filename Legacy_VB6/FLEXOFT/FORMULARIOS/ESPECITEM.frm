VERSION 5.00
Begin VB.Form ESPECITEM 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Articulos de Stock - Especificaciones"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   330
      Left            =   3570
      Picture         =   "ESPECITEM.frx":0000
      TabIndex        =   21
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   7350
      TabIndex        =   15
      Top             =   190
      Width           =   2745
      Begin VB.CommandButton Command3 
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
         Height          =   750
         Index           =   4
         Left            =   1800
         Picture         =   "ESPECITEM.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   315
         Width           =   750
      End
      Begin VB.CommandButton Command4 
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
         Height          =   750
         Left            =   1050
         Picture         =   "ESPECITEM.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Ayuda en Línea"
         Top             =   315
         Width           =   750
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   210
         Picture         =   "ESPECITEM.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   315
         Width           =   750
      End
   End
   Begin VB.Frame MARCOFICHA 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "1"
      Height          =   5895
      Index           =   1
      Left            =   735
      TabIndex        =   0
      Top             =   1365
      Width           =   10095
      Begin VB.TextBox Text28 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4740
         Left            =   735
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   735
         Width           =   8625
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Cortar"
         Height          =   260
         Left            =   4410
         TabIndex        =   7
         Top             =   5490
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Copiar"
         Height          =   260
         Left            =   5595
         TabIndex        =   6
         Top             =   5490
         Width           =   1170
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Pegar"
         Height          =   260
         Left            =   6765
         TabIndex        =   5
         Top             =   5490
         Width           =   1275
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         Height          =   260
         Left            =   8055
         TabIndex        =   4
         Top             =   5490
         Width           =   1300
      End
      Begin VB.CommandButton Command7 
         Height          =   260
         Left            =   735
         TabIndex        =   3
         Top             =   5490
         Width           =   3675
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   735
         ScaleHeight     =   375
         ScaleWidth      =   8550
         TabIndex        =   1
         Top             =   315
         Width           =   8610
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción General - Especificaciones"
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
            Left            =   105
            TabIndex        =   2
            Top             =   105
            Width           =   8415
         End
      End
      Begin VB.Line Line3 
         X1              =   9360
         X2              =   9360
         Y1              =   315
         Y2              =   5750
      End
   End
   Begin VB.Label Label2 
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
      Height          =   540
      Left            =   1470
      TabIndex        =   20
      Top             =   840
      Width           =   5625
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1470
      TabIndex        =   19
      Top             =   315
      Width           =   2115
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
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
      Left            =   270
      TabIndex        =   14
      Top             =   420
      Width           =   1005
   End
   Begin VB.Label Label3 
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
      Left            =   0
      TabIndex        =   13
      Top             =   840
      Width           =   1380
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6345
      TabIndex        =   12
      Top             =   315
      Width           =   750
   End
   Begin VB.Label Label36 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Interno:"
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
      Left            =   5250
      TabIndex        =   11
      Top             =   420
      Width           =   1005
   End
   Begin VB.Label Label37 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4560
      TabIndex        =   10
      Top             =   315
      Width           =   540
   End
   Begin VB.Label Label38 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "U-M:"
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
      Left            =   3465
      TabIndex        =   9
      Top             =   420
      Width           =   1005
   End
End
Attribute VB_Name = "ESPECITEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim YACAR%, MODIFIC%
Dim GRABATODO%

Private Sub Command1_Click()
   GRABATODO% = 1
   Call GRAESPECI
   GRABATODO% = 0
End Sub

Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.Text = Left$(Text28.Text, COMIE) + Mid$(Text28.Text, COMIE + LOSEL + 1)
End Sub

Private Sub Command3_Click(Index As Integer)
   Call GRAESPECI
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command5_Click()
   Text28.SelStart = 0
   Text28.SelLength = Len(Text28.Text)
   Text28.SetFocus
End Sub

Private Sub Command6_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.Text = Left$(Text28.Text, COMIE) + Clipboard.GetText + Mid$(Text28.Text, COMIE + LOSEL + 1)
End Sub

Private Sub Command8_Click()
   Clipboard.SetText Mid$(Text28.Text, Text28.SelStart + 1, Text28.SelLength)
End Sub

Private Sub Command9_Click()
   '
   Call GRAESPECI
   '
   Call SELECHO("MASTER")
   CODISEL$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     CODXX$ = CODISEL$
     CIXXI% = CODINT%(CODISEL$)
     Label1.Caption = VALACT1$("MASTER")
     Label2.Caption = DESCRIT0$(CIXXI%)
     Label37.Caption = UNIMED$(CIXXI%)
     Label28.Caption = Str$(CIXXI%)
   End If
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call GRAESPECI
   Call CIERRARCH("*.*")
   YACAR% = 0
End Sub
'
Private Sub Label1_Change()
   YACAR% = 0
   Call CARESPECI
End Sub

Sub GRAESPECI()
   '
   CI% = CODINT%(Label1.Caption)
   If CI% < 1 Or CI% > NUMAS% Then Exit Sub
   If YACAR% < 1 Then Exit Sub
   If MODIFIC% < 1 Then Exit Sub
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se ha Modificado la Especificación del Artículo\'" + TRIM$(CODEXT$(CI%)) + " - " + TRIM$(Label2.Caption) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR% = 0
       MODIFIC% = 0
       Call CARESPECI
       Exit Sub
     End If
   End If
   '
   Screen.MousePointer = 11
   URN& = ULTREG&("ESPECIFI")
   RNC$ = RECFILT$("ESPECIFI", 1, MKI$(CI%))
   KK& = (-3)
   '
   TEXTO$ = Text28.Text + Chr$(13) + Chr$(10)
   Call FRATEXTO(TEXTO$, 62)
   For LL% = 1 To CARETEX%
      TTYY$ = MKI$(CI%) + RETEX$(LL%)
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG("ESPECIFI", TTYY$, RENOTA&)
   Next LL%
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("ESPECIFI", MKI$(0) + Space$(62), RENOTA&)
   Wend
   '
   Call CIERRARCH("ESPECIFI")
   MODIFIC% = 0
   Screen.MousePointer = 1
   '
End Sub
'

Sub CARESPECI()
   '
   If YACAR% < 1 Then
      '
      CODD$ = Label1.Caption
      CI% = CODINT%(CODD$)
      If CI% < 1 Then Exit Sub
      '
      Text28.Text = ""
      If CI% > 0 Then
         If CI% <= NUMAS% Then
            Screen.MousePointer = 11
            RNC$ = RECFILT$("ESPECIFI", 1, MKI$(CI%))
            URN& = ULTREG&("ESPECIFI")
            '
            For U& = 1 To Len(RNC$) Step 4
               RENOTA& = CVS(Mid$(RNC$, U&, 4))
               If RENOTA& > 0 Then
                  If RENOTA& <= URN& Then
                     TTYY$ = RTrim$(Mid$(REGLEIDO$("ESPECIFI", RENOTA&), 3))
                     If Text28.Text <> "" Then
                        Text28.Text = Text28.Text + Chr$(13) + Chr$(10)
                     End If
                     Text28.Text = Text28.Text + RTrim$(TTYY$)
                  End If
               End If
            Next U&
            Screen.MousePointer = 1
         End If
      End If
      Text28.Refresh
      '
      YACAR% = 1
      MODIFIC% = 0
      '
   End If
   '
End Sub
'

Private Sub Label1_DblClick()
   '
   Call GRAESPECI
   '
   Call SELECHO("MASTER")
   CODISEL$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     CODXX$ = CODISEL$
     CIXXI% = CODINT%(CODISEL$)
     Label1.Caption = VALACT1$("MASTER")
     Label2.Caption = DESCRIT0$(CIXXI%)
     Label37.Caption = UNIMED$(CIXXI%)
     Label28.Caption = Str$(CIXXI%)
   End If
   '
End Sub

Private Sub Text28_Change()
   MODIFIC% = 1
End Sub
