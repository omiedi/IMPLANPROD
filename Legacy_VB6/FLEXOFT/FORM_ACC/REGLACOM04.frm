VERSION 5.00
Begin VB.Form REGLACOM04 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reglamentos de Compra"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Reglamento - Encabezado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8040
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Preparó:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Index           =   0
         Left            =   3360
         TabIndex        =   13
         Top             =   2310
         Width           =   4845
      End
      Begin VB.TextBox Text3 
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
         Height          =   300
         Left            =   2100
         TabIndex        =   12
         Top             =   380
         Width           =   5820
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   11
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label Label37 
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
         Height          =   300
         Left            =   1260
         TabIndex        =   10
         Top             =   375
         Width           =   720
      End
   End
   Begin VB.Frame MARCOFICHA 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Reglamento - Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6675
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1050
      Width           =   10425
      Begin VB.CommandButton Command7 
         Height          =   260
         Left            =   105
         TabIndex        =   4
         Top             =   6270
         Width           =   5355
      End
      Begin VB.TextBox Text28 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5520
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   735
         Width           =   10170
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Cortar"
         Height          =   260
         Left            =   5460
         TabIndex        =   8
         Top             =   6270
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Copiar"
         Height          =   260
         Left            =   6585
         TabIndex        =   7
         Top             =   6270
         Width           =   1170
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Pegar"
         Height          =   260
         Left            =   7740
         TabIndex        =   6
         Top             =   6270
         Width           =   1275
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         Height          =   260
         Left            =   9000
         TabIndex        =   5
         Top             =   6270
         Width           =   1290
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   105
         ScaleHeight     =   375
         ScaleWidth      =   10110
         TabIndex        =   2
         Top             =   315
         Width           =   10170
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Instrucciones del Reglamento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   0
            TabIndex        =   3
            Top             =   90
            Width           =   9885
         End
      End
      Begin VB.Line Line1 
         X1              =   105
         X2              =   10280
         Y1              =   6250
         Y2              =   6250
      End
      Begin VB.Line Line3 
         X1              =   10275
         X2              =   10275
         Y1              =   315
         Y2              =   6480
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   975
      Left            =   8260
      ScaleHeight     =   915
      ScaleWidth      =   4125
      TabIndex        =   14
      Top             =   0
      Width           =   4185
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
         Height          =   690
         Left            =   1590
         Picture         =   "REGLACOM04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   105
         Width           =   690
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
         Height          =   690
         Left            =   900
         Picture         =   "REGLACOM04.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   105
         Width           =   690
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "REGLACOM04.frx":0454
         Top             =   380
         Width           =   1515
      End
   End
End
Attribute VB_Name = "REGLACOM04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   Call GRAINSTRUC
   '
End Sub
Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub
Private Sub Command3_Click()
    Call CIERRARCH("*.*")
    Unload Me
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call MENSERR(24, "Ayuda en Línea no Disponible")
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Text28.SelStart = 0
   Text28.SelLength = Len(Text28.TEXT)
   Text28.SetFocus
End Sub

Private Sub Command6_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   ABC$ = Clipboard.GetText
10 For KKI% = 1 To Len(ABC)
     If Asc(Mid$(ABC$, KKI%, 1)) = 9 Then
       ABC$ = Left$(ABC$, KKI% - 1) + Space$(5) + Mid$(ABC$, KKI% + 1)
       GoTo 10
     End If
   Next KKI%
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + ABC$ + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command8_Click()
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
End Sub
Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    '
    Refresh
End Sub

Private Sub Form_Resize()
   DoEvents
   If WindowState = 1 Then
      Call BAJALDISCO
   End If
End Sub

Private Sub Label28_Change()
   '
   Call CARINSTRUC
   '
End Sub
Sub CARINSTRUC()
   '
   CIIINSTRU% = Val(Label37)
   If CIIINSTRU% < 1 Then Exit Sub
   '
   REGI& = CIIINSTRU%
   URN& = ULTREG&("REGLACOM")
   '
   Text28.TEXT = "": TTYZ$ = ""
   If CIIINSTRU% > 0 Then
     If CIIINSTRU% <= URN& Then
       Screen.MousePointer = 11
       RNC$ = RECFILT$("REGLACOM", 1, MKI$(CIIINSTRU%))
       '
       For U& = 1 To Len(RNC$) Step 4
         RENOTA& = CVS(Mid$(RNC$, U&, 4))
         If RENOTA& > 0 Then
           If RENOTA& <= URN& Then
             TTYY$ = RTrim$(Mid$(REGLEIDO$("REGLACOM", RENOTA&), 3))
             If TTYZ$ <> "" Then
               TTYZ$ = TTYZ$ + Chr$(13) + Chr$(10)
             End If
             TTYZ$ = TTYZ$ + RTrim$(TTYY$)
           End If
         End If
       Next U&
       Screen.MousePointer = 1
     End If
   End If
   '
   LOCADE& = Len(TTYZ$)
   For U& = LOCADE& To 1 Step -1
     If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 30
   Next U&
   U& = 0
   '
30 Text28.TEXT = Left$(TTYZ$, U&)
   Text28.Refresh
   '
End Sub
'
Sub GRAINSTRUC()
   '
   URN& = ULTREG&("REGLACOM")
   CIIINSTRU% = Val(Label37)
   If CIIINSTRU% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   Call BLOQARCH("REGLACOM")
   RNC$ = RECFILT$("REGLACOM", 1, MKI$(CIIINSTRU%))
   KK& = (-3)
   '
   TTYZ$ = Text28.TEXT
   LOCADE& = Len(TTYZ$)
   For U& = LOCADE& To 1 Step -1
     If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 10
   Next U&
   U& = 0
   '
10 TEXTO1$ = Left$(TTYZ$, U&) + Chr$(13) + Chr$(10)
   Call FRATEXTO(TEXTO1$, 92)
   For LL% = 1 To CARETEX%
      TTYY$ = MKI$(CIIINSTRU%) + LTrim$(RETEX$(LL%)) + Space$(128)
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG("REGLACOM", TTYY$, RENOTA&)
   Next LL%
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("REGLACOM", MKI$(0) + Space$(92), RENOTA&)
   Wend
   '
   Call CIERRARCH("REGLACOM")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Screen.MousePointer = 1
   '
End Sub
''
'Sub PRINPROCE(DOCUORIGEN$)
'  '
'  If List1.ListCount > 0 Then
'     '
'     FORIPRE$ = TRIM$(CONTROL$("FORGRAFI", "DOCUGRA4"))
'     If FORIPRE$ <> "" Then
'        Screen.MousePointer = 11
'        FEX = HOY(HOS$)
'        FECHDOC$ = HOS$
'        NUMEDOC$ = ""
'        TIPODOC$ = "Hoja de Actualización " + TRIM$(Label37)
'        '
'        CODPARAM$(1) = "COD-MAT"
'        CODPARAM$(2) = "REF-MAT2"
'        CODPARAM$(3) = "DES-MAT"
'        CODPARAM$(4) = "VENDEDOR"
'        CODPARAM$(5) = "COND-PAG"
'        CODPARAM$(6) = "FECH-VEN"
'        CAPARDOC% = 6
'        '
'        DOCPARAM$(1) = TRIM$(Label37)
''        DOCPARAM$(2) = TRIM$(Text1)
''        DOCPARAM$(3) = TRIM$(Text3)
''        DOCPARAM$(4) = TRIM$(Text4(0))
''        DOCPARAM$(5) = TRIM$(Text4(1))
''        DOCPARAM$(6) = TRIM$(Text5(1))
'        '
'        CODTABU1$(1) = "NOP-OPER"
'        CODTABU1$(2) = "DES-OPER"
'        CODTABU1$(3) = "OBS-OPER"
'        CACOLTAB1% = 3
'        '
'        CAITAB1% = 0
'        ZANT$ = ""
'        For U& = 1 To List1.ListCount
'          Z$ = List1.List(U& - 1)
'          If TRIM$(Left$(Z$, 18)) <> "" Then
'              CAITAB1% = CAITAB1% + 1
'              TETABU1$(1, CAITAB1%) = Mid$(Z$, 10, 4)
'              TETABU1$(2, CAITAB1%) = Mid$(Z$, 15, 4)
'              TETABU1$(3, CAITAB1%) = RTrim$(Mid$(Z$, 22, 62))
'            Else
'              TETABU1$(3, CAITAB1%) = TETABU1$(3, CAITAB1%) + " " + RTrim$(Mid$(Z$, 22, 62))
'          End If
'        Next U&
'        '
'        Call PRINTDOC("@" + FORIPRE$)
'        Screen.MousePointer = 1
'        '
'     End If
'  End If
'  '
'  FORIPRE$ = TRIM$(CONTROL$("FORGRAFI", "DOCUGRA2"))
'
'  If FORIPRE$ = "" Then
'       PRINTFORM
'     Else
'       Screen.MousePointer = 11
'       FEX = HOY(HOS$)
'       FECHDOC$ = HOS$
'       NUMEDOC$ = ""
'       TIPODOC$ = "Procedimiento " + TRIM$(Label37)
'       '
'       CODPARAM$(1) = "COD-MAT"
'       CODPARAM$(2) = "REF-MAT2"
'       CODPARAM$(3) = "DES-MAT"
'       CODPARAM$(4) = "VENDEDOR"
'       CODPARAM$(5) = "COND-PAG"
'       CODPARAM$(6) = "FECH-VEN"
'       CAPARDOC% = 6
'       '
'       DOCPARAM$(1) = TRIM$(Label37)
''       DOCPARAM$(2) = TRIM$(Text1)
''       DOCPARAM$(3) = TRIM$(Text3)
''       DOCPARAM$(4) = TRIM$(Text4(0))
''       DOCPARAM$(5) = TRIM$(Text4(1))
''       DOCPARAM$(6) = TRIM$(Text5(1))
'       '
'       CODTABU1$(1) = "REF-MAT1"
'       CACOLTAB1% = 1
'       '
'       CAITAB1% = 0
'       Call FRATEXTO(Text28, 92)
'       For U& = 1 To CARETEX%
'         Z$ = "." + RETEX$(U&)
'         CAITAB1% = CAITAB1% + 1
'         TETABU1$(1, CAITAB1%) = Z$
'       Next U&
'       '
'       Call PRINTDOC("@" + FORIPRE$)
'       Screen.MousePointer = 1
'       '
'  End If
'  '
'End Sub

Private Sub Label37_CHANGE()
  '
  Text3 = ECOARCH$("INDIRCOM", MKI$(Val(Label37)))
  Call CARINSTRUC
  '
End Sub

Private Sub text3_gotfocus()
  '
  Text28.SetFocus
  '
End Sub
