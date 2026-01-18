VERSION 5.00
Begin VB.Form GENERDEMO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Genera CD de Demostración"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
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
      TabIndex        =   9
      Top             =   2415
      Width           =   1590
   End
   Begin VB.TextBox Text2 
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
      Left            =   315
      TabIndex        =   8
      Text            =   "F:\FLEXOFT\NOVEDAD"
      Top             =   1470
      Width           =   5055
   End
   Begin VB.TextBox Text1 
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
      Left            =   315
      TabIndex        =   6
      Text            =   "C:\CD_DEMO\FX_PROGRAMS"
      Top             =   525
      Width           =   5055
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   2940
      ScaleHeight     =   75
      ScaleWidth      =   750
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3465
      Width           =   810
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   750
      Left            =   4515
      Picture         =   "GENERDEMO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Inicia Inoculación"
      Top             =   2900
      Width           =   855
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   315
      TabIndex        =   0
      Top             =   1995
      Width           =   2430
   End
   Begin VB.Label Label6 
      Caption         =   "Fecha de Vencimiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2835
      TabIndex        =   10
      Top             =   2100
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Ruta a Novedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      TabIndex        =   7
      Top             =   1155
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Ruta a Programas CD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      TabIndex        =   5
      Top             =   210
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "GENERDEMO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   '
   DoEvents
   Screen.MousePointer = 11
   '
   HOII% = HOY(HOS$)
   FEMODUI$ = TRIM$(Str$(HOII%))
   While Len(FEMODUI$) > 4: FEMODUI$ = Mid$(FEMODUI$, 2): Wend
   FEVENCI$ = TRIM$(Str$(FECHANUM(Text3)))
   While Len(FEVENCI$) > 4: FEVENCI$ = Mid$(FEVENCI$, 2): Wend
   '
   FIN& = File1.ListCount
   RUTACD$ = Text1
   RUTANOVE$ = Text2
   '
   For U& = 1 To File1.ListCount
     Call TRACE(20, U&, FIN&)
     FINAX$ = UCase$(File1.List(U& - 1))
 If InStr(FINAX$, "FLOGON") > 0 Then
 AAA = BBB
 End If
     If EXISTE%(RUTANOVE$ + "\" + FINAX$) > 0 Then
       '
       FINAJ$ = AJUSTI$(FINAX$, 8)
       FEMODUJ$ = ""
       For KKI% = 1 To 4
         CARAJ$ = TRIM$(Str$((Val(Mid$(FEMODUI$, KKI%, 1)) + Asc(Mid$(FINAJ$, KKI%, 1))) Mod 10))
         FEMODUJ$ = FEMODUJ$ + CARAJ$
       Next KKI%
       FEVENCJ$ = ""
       For KKI% = 1 To 4
         CARAJ$ = TRIM$(Str$((Val(Mid$(FEVENCI$, KKI%, 1)) + Asc(Mid$(FINAJ$, KKI% + 4, 1))) Mod 10))
         FEVENCJ$ = FEVENCJ$ + CARAJ$
       Next KKI%
       '
       ARCHORI$ = LOADFILE$(RUTANOVE$ + "\" + FINAX$)
       IKYY$ = Space$(9)
       Call REPLA(IKYY$, FEMODUJ$, 1, 4)
       Call REPLA(IKYY$, FEVENCJ$, 6, 4)
       ' DIGITO VERIFICADOR DEL SECTOR DE CONTROL
       DIVERI& = 0
       For KKI% = 1 To 8
         KKJ% = KKI%: If KKI% > 4 Then KKJ% = KKI% + 1
         DIVERI& = Val(Mid$(IKYY$, KKJ%, 1)) + (7 * KKJ% ^ 2)
       Next KKI%
       DIVERI& = DIVERI& Mod 10
       DIVERX$ = TRIM$(Str$(DIVERI&))
       Call REPLA(IKYY$, DIVERX$, 5, 1)
       '
       PQXX& = InStr(ARCHORI$, "Microsoft Vísual Basic 5.O")
       If PQXX& > 0 Then
         If Mid$(ARCHORI$, PQXX& + 50, 9) = "........." Then
           Mid$(ARCHORI$, PQXX& + 50, 9) = IKYY$
         End If
         DESTI$ = RUTACD$ + "\" + FINAX$
         Call SAVEFILE(DESTI$, ARCHORI$)
       End If
       '
     End If
   Next U&
   '
   '
   Screen.MousePointer = 1
   Close #NX%
   '
   MsgBox "Proceso Completo"
   '
End Sub


Private Sub Form_Load()
   File1.Path = Text1
   File1.Pattern = "*.exe"
   HOII% = HOY(HO$)
   Text3 = HO$
End Sub

Private Sub Text1_Change()
   On Error Resume Next
   File1.Path = Text1
   File1.Pattern = "*.exe"
   On Error GoTo 0
End Sub

