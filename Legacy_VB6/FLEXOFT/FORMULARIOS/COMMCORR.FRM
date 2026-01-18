VERSION 5.00
Begin VB.Form COMMCORR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comentarios y Correcciones"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "COMMCORR.frx":0000
      Top             =   960
      Width           =   11175
   End
   Begin VB.CommandButton Command1 
      Height          =   700
      Left            =   10440
      Picture         =   "COMMCORR.frx":0006
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Inicia Inoculación"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
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
      Height          =   330
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Tecnico:"
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
      Left            =   6360
      TabIndex        =   5
      Top             =   240
      Width           =   3060
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
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
      Height          =   330
      Left            =   9000
      TabIndex        =   4
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Rutina:"
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
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   3060
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
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
      Height          =   330
      Left            =   4440
      TabIndex        =   2
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Módulo:"
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
      Left            =   -2160
      TabIndex        =   1
      Top             =   225
      Width           =   3060
   End
End
Attribute VB_Name = "COMMCORR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CARGACORREC()
   CONDI$ = "TECRESPO = '" & Label5 & "' "
   CONDI$ = CONDI$ & " AND MODULO = '" & TRIM$(Left$(Label2, 64)) & "' "
   CONDI$ = CONDI$ & " AND RUTINA = '" & Label3 & "'"
   Text1 = VALOBADA("CORRPROG", "Deta_corr", CONDI$, "NOMESS")
End Sub

Private Sub Command1_Click()
   DTC$ = Text1
   If Left$(TRIM$(Right$(DTC$, 16)), 12) <> "============" Then
      DTC$ = DTC$ & vbCrLf & String$(100, "=")
   End If
   '
   If TRIM$(Left$(Label2, 64)) <> "" Then
     If TRIM$(Label3) <> "" Then
       If TRIM$(Label5) <> "" Then
         CONDI$ = "TECRESPO = '" & Label5 & "' "
         CONDI$ = CONDI$ & " AND MODULO = '" & TRIM$(Left$(Label2, 64)) & "' "
         CONDI$ = CONDI$ & " AND RUTINA = '" & Label3 & "'"
         QTR& = CantRegistros("CORRPROG", CONDI$)
         If QTR& < 1 Then
             Call CreaRegistro("CORRPROG", "TECRESPO", "NUEVACORREC")
             CONDI2$ = "TECRESPO = 'NUEVACORREC'"
             Call ACTUREGISTRO("CORRPROG", "MODULO", CONDI2$, TRIM$(Left$(Label2, 64)), RAFE&, "NOMESS")
             Call ACTUREGISTRO("CORRPROG", "RUTINA", CONDI2$, Label3, RAFE&, "NOMESS")
             Call ACTUREGISTRO("CORRPROG", "TECRESPO", CONDI2$, Label5, RAFE&, "NOMESS")
         End If
         Call ACTUREGISTRO("CORRPROG", "Deta_Corr", CONDI$, DTC$, RAFE&, "NOMESS")
         Unload Me
       End If
     End If
   End If
End Sub

Private Sub Form_Load()
   ' ESTE ES UN COMENTARIO
End Sub
