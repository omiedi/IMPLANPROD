VERSION 5.00
Begin VB.Form ANACOSMOD 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis de Costos de Mano de Obra"
   ClientHeight    =   6330
   ClientLeft      =   135
   ClientTop       =   2070
   ClientWidth     =   11715
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Height          =   330
      Left            =   840
      Picture         =   "ANACOSMOD.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Listado (Tabla) de Costo de Mano de Obra"
      Top             =   5880
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      Height          =   330
      Left            =   210
      Picture         =   "ANACOSMOD.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   540
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5460
      Left            =   -30
      TabIndex        =   2
      Top             =   315
      Width           =   11755
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12000
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   $"ANACOSMOD.frx":024C
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   330
         Left            =   40
         TabIndex        =   4
         Top             =   50
         Width           =   11880
      End
   End
   Begin VB.Line Line2 
      X1              =   11320
      X2              =   11320
      Y1              =   5900
      Y2              =   6200
   End
   Begin VB.Line Line1 
      X1              =   5880
      X2              =   11325
      Y1              =   6195
      Y2              =   6195
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
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
      Height          =   330
      Left            =   10120
      TabIndex        =   7
      Top             =   5880
      Width           =   1210
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
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
      Height          =   330
      Left            =   8925
      TabIndex        =   6
      Top             =   5880
      Width           =   1210
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Height          =   330
      Left            =   7730
      TabIndex        =   0
      Top             =   5880
      Width           =   1210
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      Height          =   330
      Left            =   5895
      TabIndex        =   5
      Top             =   5880
      Width           =   1845
   End
   Begin VB.Line Line9 
      X1              =   0
      X2              =   11865
      Y1              =   5775
      Y2              =   5775
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Totales:"
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
      Left            =   3045
      TabIndex        =   1
      Top             =   5955
      Width           =   2745
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   10500
      Y1              =   5060
      Y2              =   5060
   End
End
Attribute VB_Name = "ANACOSMOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   On Error Resume Next
   ANACOSTO.COMMLAB.Caption = "PRIMOD"
   ANACOSTO04.COMMLAB.Caption = "PRIMOD"
   On Error GoTo 0
End Sub

Private Sub Command2_Click()
 '\\\OT-06-0161-WAR-28/05/06///
  Command2.Enabled = False
  Screen.MousePointer = 11
  JJ& = 0
  Y$ = REGBLAN$("LICOSMOD")
  For j% = 0 To List1.ListCount - 1
    X$ = List1.List(j%)
    '
    CE1$ = Mid$(X$, 1, 15)
    CIA% = CI%
    CI% = CODINT%(CE1$)
    If TRIM$(CE1$) = "" And CI% = 0 Then
     CI% = CIA%
    End If
    OPERA1$ = Mid$(X$, 39, 15)
    If OPERA1$ = "---------------" Or OPERA1$ = "===============" Then
      GoTo 10
    End If
    CAOP1 = Val(Mid$(X$, 199, 5))
    MINUT1$ = Mid$(X$, 55, 10)
    '
    SEG1 = Val(Right$(MINUT1$, 3))
    HOR1 = Val(Left$(MINUT1$, 3))
    MINUT2 = Val(Mid$(MINUT1$, 5, 2))
    '
    SEG2 = SEG1 / 60 'PASO SEGUNDOS A MINUTOS
    HOR2 = HOR1 * 60 'PASO HORAS A MINUTOS
    MINUTOT1 = SEG2 + HOR2 + MINUT2
    '
    PUNPES1# = Val(Mid$(X$, 204, 10))
    CMOD1# = Val(Mid$(X$, 65, 10))
    CEQUI1# = Val(Mid$(X$, 75, 10))
    CTOT1# = Val(Mid$(X$, 85, 10))
    '
    Call REPLA(Y$, MKI$(CI%), 1, 2)
    Call REPLA(Y$, OPERA1$, 3, 19)
    Call REPLA(Y$, MKS$(CAOP1), 21, 4)
    Call REPLA(Y$, MKS$(MINUTOT1), 25, 4)
    Call REPLA(Y$, MKD$(PUNPES1#), 29, 8)
    Call REPLA(Y$, MKD$(CMOD1#), 37, 8)
    Call REPLA(Y$, MKD$(CEQUI1#), 45, 8)
    Call REPLA(Y$, MKD$(CTOT1#), 53, 8)
    JJ& = JJ& + 1
    Call GRAREG("LICOSMOD", Y$, JJ&)
    '
10 Next j%
  Screen.MousePointer = 1
  Call SETULTREG("LICOSMOD", JJ&)
  Call HEADERS("LICOSMOD", "")
  On Error Resume Next
     TTT1$ = ANACOSTO.Text1 + " - " + ANACOSTO.Text3
     If Err Then
        On Error GoTo 0
        TTT1$ = ANACOSTO04.Text1 + " - " + ANACOSTO04.Text3
     End If
  On Error GoTo 0
  '
  Call HEADERS("LICOSMOD", "Articulo: " + TTT1$)
  Call FINAL("*LICOSMOD")
  Command2.Enabled = True
  '\\\OT-06-0161-WAR-FIN///
End Sub

Private Sub Form_Load()
   If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
   End If

   Call CENTRAFORM(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 1
   Hide
End Sub

Private Sub List1_DblClick()
   '
   CODICOM$ = TRIM$(Left$(List1.TEXT, 16))
   If CODICOM$ = String$(16, "-") Or CODICOM$ = String$(16, "=") Then
     Exit Sub
   End If
   '
   On Error Resume Next
   TTXX1$ = ""
   TTXX1$ = TTXX1$ + ANACOSTO.Text1
   TTXX1$ = TTXX1$ + ANACOSTO04.Text1
   On Error GoTo 0
   '
   If TRIM$(TTXX1$) <> "" Then
     LLAMACOSTO$ = TTXX1$
     ORILLAMA$ = "MOD"
     REGILLAMA% = List1.ListIndex
     TOPELLAMA% = List1.TopIndex
   End If
   '
   COPER% = 0
   For KKK% = List1.ListIndex To 0 Step -1
     COD$ = TRIM$(Left$(List1.List(KKK%), 16))
     If COD$ <> "" Then
       If COD$ <> String$(16, "-") Then
         On Error Resume Next
           ANACOSTO.SSTab1.Tab = 2
           ANACOSTO04.SSTab1.Tab = 2
           DoEvents
           ANACOSTO.Text1.TEXT = COD$
           ANACOSTO04.Text1.TEXT = COD$
           DoEvents
           ANACOSTO.List3.ListIndex = COPER%
           ANACOSTO04.List3.ListIndex = COPER%
         On Error GoTo 0
         Hide
         Exit Sub
       End If
     End If
     COPER% = COPER% + 1
   Next KKK%

End Sub
