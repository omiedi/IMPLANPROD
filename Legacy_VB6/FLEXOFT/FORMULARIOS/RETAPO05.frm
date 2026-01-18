VERSION 5.00
Begin VB.Form RETAPO05 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Retenciones legales y Aportes"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   225
      ScaleHeight     =   195
      ScaleWidth      =   5475
      TabIndex        =   2
      Top             =   170
      Width           =   5535
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código  Descripción"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   20
         Width           =   5535
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   4815
      Left            =   5985
      ScaleHeight     =   4755
      ScaleWidth      =   1230
      TabIndex        =   1
      Top             =   -240
      Width           =   1290
      Begin VB.CommandButton Command21 
         Caption         =   "New"
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
         Left            =   90
         Picture         =   "RETAPO05.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Alta de Nuevo Item de Retenciones legales y Aportes"
         Top             =   1980
         Width           =   650
      End
      Begin VB.CommandButton Command4 
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
         Left            =   90
         Picture         =   "RETAPO05.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Ayuda en Línea"
         Top             =   1050
         Width           =   650
      End
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
         Left            =   90
         Picture         =   "RETAPO05.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Exit"
         Top             =   360
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -330
         Picture         =   "RETAPO05.frx":059E
         Top             =   4020
         Width           =   1515
      End
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
      Left            =   225
      TabIndex        =   0
      Top             =   420
      Width           =   5535
   End
End
Attribute VB_Name = "RETAPO05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Sub Command21_Click()
 '
  Command21.Enabled = False
  'VDEF$ TIENE QUE DAR EL PROX. Nº DE REMUNERACIÓN
  VDEF$ = Space(32)
  PRONU% = Int(ULTREG&("RETLEG"))
  Call REPLA(VDEF$, MKI$(PRONU% + 1), 1, 2)
9 OBJ$ = OBJPLA$("NUEREM", VDEF$)
  If OBJ$ = "" Then GoTo 99
  VDEF$ = OBJ$
  '
  DESCRI$ = TRIM$(Mid$(OBJ$, 3, 30))
  If TRIM$(DESCRI$) = "" Then
    Call MENSERR(24, "Descripción no Válida.")
    GoTo 9
  End If
  '
  ZZ$ = Space(64)
  NUME = CVI(Mid$(OBJ$, 1, 2))
  NUMER$ = FORMATNUM$(NUME, "F5")
  NUME1$ = AJUSTD$(Val(NUMER$), 7)
  Call REPLA(ZZ$, NUME1$, 1, 7)
  Call REPLA(ZZ$, DESCRI$, 10, 30)
  List1.AddItem ZZ$
  '
  ZZ1$ = Space(128)
  Call REPLA(ZZ1$, DESCRI$, 1, 30) 'DESCRIPCION
  Call REPLA(ZZ1$, Chr$(1), 31, 1) 'ESTADO ACTIVO
  Call REPLA(ZZ1$, Chr$(1), 32, 1) 'TIPO NORMAL
  Call REPLA(ZZ1$, Chr$(0), 33, 1) 'MODO CALCULO INTERNO
  Call REPLA(ZZ1$, Chr$(0), 34, 1) 'CARACTER OPCIONAL
  Call REPLA(ZZ1$, Chr$(1), 35, 1) 'RETENCION GRAVADO
  Call REPLA(ZZ1$, Chr$(0), 36, 1) 'QUINCENA TODAS
  Call REPLA(ZZ1$, Chr$(0), 37, 1) 'RECALCULO NO
  Call REPLA(ZZ1$, Chr$(0), 38, 1) 'FILLER
  Call REPLA(ZZ1$, " ", 39, 10) 'FILLER 1
  Call REPLA(ZZ1$, " ", 49, 72) 'FORMULA
  Call REPLA(ZZ1$, " ", 121, 8) 'ABREVIATURA
  Call REGAPP("RETLEG", ZZ1$)
  '
99 Command21.Enabled = True
  '
End Sub

Private Sub Command3_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

 Sub Form_Load()
  '
  Call CENTRAFORM(Me)
  '
  List1.Clear
  For J& = 1 To ULTREG&("RETLEG")
    X$ = REGLEIDO$("RETLEG", J&)
    DESCRI$ = Mid$(X$, 1, 30)
    ZZ$ = Space(64)
    NUM$ = AJUSTD$(CStr(J&), 7)
    Call REPLA(ZZ$, NUM$, 1, 7)
    Call REPLA(ZZ$, DESCRI$, 10, 30)
    List1.AddItem ZZ$
  Next J&
  '
End Sub

Private Sub List1_DBLClick()
  '
  X$ = List1.List(List1.ListIndex)
  COD$ = Mid$(X$, 1, 7)
  '
  JJ& = Val(COD$)
  ZZ$ = REGLEIDO$("RETLEG", JJ&)
  DESCRI$ = TRIM$(Mid$(ZZ$, 1, 30))
  ABRE$ = TRIM$(Mid$(ZZ$, 121, 8))
  PROP$ = Mid$(ZZ$, 31, 8)
  FORMU$ = Mid$(ZZ$, 49, 72)
  JJ1 = CVS(Mid$(ZZ$, 39, 4)) '%RETENCION
  JJ2 = CVS(Mid$(ZZ$, 43, 4)) '%APORTES
  With RETDET05
    .Text8 = ABRE$ 'ABREVIATURA
    .Text7 = FORMU$ 'FORMULA
    EST% = Asc(Mid$(PROP$, 1, 1))
    ET$ = ESTADO(EST%) 'DEVUELVE EL VALOR EN $ DEL ESTADO
    .Text16 = ET$ 'ESTADO
    .Text1 = Asc(Mid$(PROP$, 3, 1)) 'CARACTER
    .Text2 = Asc(Mid$(PROP$, 4, 1)) 'MODO DE CALCULO
    .Text3 = Asc(Mid$(PROP$, 6, 1)) 'QUINCENA
    .Text4 = Asc(Mid$(PROP$, 7, 1)) 'RECALCULO
    If Val(.Text2) <> 3 Then 'SI MODO <> SUMAS FIJAS
      .Text5 = FORMATNUM(JJ1, "F5.2") '%RET
      .Text6 = FORMATNUM(JJ2, "F5.2") '%APORTES
    Else
      .Text11 = FORMATNUM(JJ1, "F9.2") 'IMP. RET
      .Text10 = FORMATNUM(JJ2, "F9.2") 'IMP. APORTES
    End If
    .Label15 = COD$
    .Text9 = DESCRI$
    .Show 1
  End With
  '
End Sub
Function ESTADO$(EST%) 'DEVUELVE EL VALOR EN $ DEL ESTADO
  Select Case EST%
    Case 0
      ESTADO$ = "S" 'SUSPENDIDO
    Case 1
      ESTADO$ = "A" 'ACTIVO
    Case 2
      ESTADO$ = "F" 'FANTASMA
  End Select
End Function
