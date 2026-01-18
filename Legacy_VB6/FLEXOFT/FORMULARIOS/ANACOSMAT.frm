VERSION 5.00
Begin VB.Form ANACOSMAT 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis de Costos de Materiales"
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
      Picture         =   "ANACOSMAT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Listado (Tabla) de Costo de Suministros"
      Top             =   5880
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      Height          =   330
      Left            =   210
      Picture         =   "ANACOSMAT.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ficha sw Análisis de Costo de Suministros"
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
      Left            =   0
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
         Caption         =   $"ANACOSMAT.frx":024C
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
      X1              =   6520
      X2              =   11330
      Y1              =   6200
      Y2              =   6200
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
      TabIndex        =   6
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
      Left            =   8925
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
      Left            =   7730
      TabIndex        =   5
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
      Left            =   6530
      TabIndex        =   0
      Top             =   5880
      Width           =   1210
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
      Left            =   3550
      TabIndex        =   1
      Top             =   5950
      Width           =   2745
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   10500
      Y1              =   5060
      Y2              =   5060
   End
End
Attribute VB_Name = "ANACOSMAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   On Error Resume Next
   ANACOSTO.COMMLAB.Caption = "PRISUMI"
   ANACOSTO04.COMMLAB.Caption = "PRISUMI"
   On Error GoTo 0
End Sub

Private Sub Command2_Click()
  '\\\OT-06-0133-WAR-06/05/06///
  Command2.Enabled = False
  Screen.MousePointer = 11
  JJ& = 0
  Y$ = REGBLAN$("LICOSMAT")
  For j% = 0 To List1.ListCount - 1
    x$ = List1.List(j%)
    '
    CE1$ = Mid$(x$, 1, 15)
    CI% = CODINT%(CE1$)
    Q1 = Val(Mid$(x$, 40, 10))
    UM1$ = Mid$(x$, 52, 3)
    PUN1# = Val(Mid$(x$, 169, 10))
    MON1$ = Mid$(x$, 179, 10)
    PUND1# = Val(Mid$(x$, 189, 10))
    CMAPRI1# = Val(Mid$(x$, 55, 10))
    CPIE1# = Val(Mid$(x$, 65, 10))
    CTER1# = Val(Mid$(x$, 75, 10))
    CTOT1# = Val(Mid$(x$, 85, 10))
    '
    Call REPLA(Y$, MKI$(CI%), 1, 2)
    Call REPLA(Y$, MKS$(Q1), 3, 4)
    Call REPLA(Y$, UM1$, 7, 4)
    Call REPLA(Y$, MKD$(PUN1#), 11, 8)
    Call REPLA(Y$, MON1$, 19, 10)
    Call REPLA(Y$, MKD$(PUND1#), 29, 8)
    Call REPLA(Y$, MKD$(CMAPRI1#), 37, 8)
    Call REPLA(Y$, MKD$(CPIE1#), 45, 8)
    Call REPLA(Y$, MKD$(CTER1#), 53, 8)
    Call REPLA(Y$, MKD$(CTOT1#), 61, 8)
    JJ& = JJ& + 1
    Call GRAREG("LICOSMAT", Y$, JJ&)
    '
  Next j%
  Screen.MousePointer = 1
  Call SETULTREG("LICOSMAT", JJ&)
  Call HEADERS("LICOSMAT", "")
  On Error Resume Next
     TTT1$ = ANACOSTO.Text1 + " - " + ANACOSTO.Text3
     If Err Then
        On Error GoTo 0
        TTT1$ = ANACOSTO04.Text1 + " - " + ANACOSTO04.Text3
     End If
  On Error GoTo 0
  '
  Call HEADERS("LICOSMAT", "Articulo: " + TTT1$)
  Call FINAL("*LICOSMAT")
  Command2.Enabled = True
  '\\\OT-06-0133-WAR-FIN///
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
     ORILLAMA$ = "MAT"
     REGILLAMA% = List1.ListIndex
     TOPELLAMA% = List1.TopIndex
   End If
   '
   COD$ = TRIM$(Left$(List1.TEXT, 16))
   If COD$ <> "" Then
     Hide
     On Error Resume Next
     ANACOSTO.SSTab1.Tab = 0
     ANACOSTO04.SSTab1.Tab = 0
     DoEvents
     ANACOSTO.Text1.TEXT = COD$
     ANACOSTO04.Text1.TEXT = COD$
     On Error GoTo 0
   End If
   '
End Sub

