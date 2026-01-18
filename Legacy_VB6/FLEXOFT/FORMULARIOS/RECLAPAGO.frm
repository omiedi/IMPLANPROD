VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form RECLAPAGO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aviso de Vencimiento"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   16
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar Mensaje"
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
      Left            =   3045
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7350
      Width           =   1905
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Recibo Electrónico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   315
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Baja Recibo Electronico de la WEB"
      Top             =   3675
      Width           =   2115
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
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
      Height          =   1320
      Left            =   210
      TabIndex        =   4
      Top             =   2205
      Width           =   7365
   End
   Begin VB.PictureBox Picture1 
      Height          =   330
      Left            =   210
      ScaleHeight     =   270
      ScaleWidth      =   7305
      TabIndex        =   2
      Top             =   1890
      Width           =   7365
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto            F.Venc.  M.  Importe  Interés    Total"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   50
         TabIndex        =   3
         Top             =   80
         Width           =   7365
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "RECLAPAGO.frx":0000
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "RECLAPAGO.frx":007A
      Top             =   1200
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "(Más I.V.A.)"
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
      Left            =   5565
      TabIndex        =   14
      Top             =   4095
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "0"
      Height          =   330
      Left            =   105
      TabIndex        =   13
      Top             =   630
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   $"RECLAPAGO.frx":02AE
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   315
      TabIndex        =   1
      Top             =   1155
      Width           =   7155
   End
   Begin VB.Label Label10 
      Caption         =   "De nuestra mayor consideración:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   315
      TabIndex        =   12
      Top             =   900
      Width           =   7155
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2310
      TabIndex        =   11
      Top             =   210
      Width           =   5055
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
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
      TabIndex        =   10
      Top             =   3780
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   5565
      TabIndex        =   9
      Top             =   3675
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "FLEXOFT Soluciones"
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
      Left            =   4830
      TabIndex        =   8
      Top             =   6930
      Width           =   2430
   End
   Begin VB.Label Label5 
      Caption         =   "Atentamente,"
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
      TabIndex        =   7
      Top             =   6510
      Width           =   6210
   End
   Begin VB.Label Label4 
      Caption         =   $"RECLAPAGO.frx":0338
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   315
      TabIndex        =   6
      Top             =   5565
      Width           =   7365
   End
   Begin VB.Label Label3 
      Caption         =   $"RECLAPAGO.frx":03DB
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   315
      TabIndex        =   5
      Top             =   4620
      Width           =   7260
   End
End
Attribute VB_Name = "RECLAPAGO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim quieto%
Dim demora
Dim VECES

Private Sub Command1_Click()
   quieto% = 0: VECES = 0
   Unload Me
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If quieto% <> 1 Then
     Randomize Timer
     x = (Width - Command1.Width) * Rnd
     Y = (Height - Command1.Height - 300) * Rnd
     Command1.Left = x
     Command1.Top = Y
   End If
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   SVLF$ = SEVOLF
   If SVLF$ <> "" Then
      Screen.MousePointer = 11
      ARCHA$ = SVLF$ + ".REL"
      JJ1& = BAJADELAWEB(ARCHA$)
      If JJ1& <> 0 Then
         Call MENSERR(24, "Error " & JJ1& & "\Imposible Bajar " & ARCHA$ & "\Correspondiente al Recibo Electrónico.\  \Comuniquese con FLEXOFT. Muchas Gracias")
         GoTo 99
      End If
      '
      SQLX$ = "SELECT * FROM CONTRO_DOWN"
      SQLX$ = SQLX$ + " WHERE File_Name = '" & ARCHA$ & "'"
      '
      On Error Resume Next
      Set MIRS = BADACON.OpenRecordset(SQLX$, dbOpenDynaset)
      MIRS.MoveFirst
      If Err Then
         On Error GoTo 0
         GoTo 99                 ' no hay nada que bajar
      End If
      On Error GoTo 0
      '
      With MIRS
         .Edit
            AHORA1$ = Format$(Now, "dd/mm/yy  hh:mm")
         !FeHor_Baja = AHORA1$
         !Usu_Baja = USERNA$
         .Update
      End With
      MIRS.Close
      Screen.MousePointer = 1
      Call COMUNI("Recibo Electrónico Bajado.\\Re-Inicie Ahora FLEXOFT para Instalar.")
   End If
   '
99 Screen.MousePointer = 1
   Command2.Enabled = True
End Sub

Private Sub Form_Load()
   demora = 10
   '
   CD$ = Left$(App.Path, 2)
   DoEvents
   If Mid$(CD$, 2, 1) = ":" Then
      SEVOLF$ = TRIM$(Left$(SERIVOL$(CD$) + Space$(12), 12))
   End If
   Command2.Caption = Command2.Caption + " (" + SEVOLF$ + ")"
   Timer1.Enabled = True
    Call APLICACIONSKINS(Me)
End Sub

Private Sub Timer1_Timer()
   '
   If quieto% = 1 Then Exit Sub
   QTDEM = (Val(Label11) - 25)
   If QTDEM < 1 Then
     quieto% = 1
     Exit Sub
   End If
   '
   TDEMOR = 1
   For KKUU& = 1 To QTDEM
     TDEMOR = TDEMOR * 1.12
   Next KKUU&
   If TDEMOR > 300 Then TDEMOR = 300
   '
'Label13 = Label11 & "   " & TDEMOR & "   " & VECES
'Label13.Refresh
   VECES = VECES + 1
   If VECES > TDEMOR Then quieto% = 1
End Sub
