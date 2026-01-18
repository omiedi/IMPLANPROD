VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form NUEMPRE09 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nueva Empresa"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   7
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   120
      Top             =   1185
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00004040&
      Height          =   8600
      Left            =   6705
      ScaleHeight     =   8535
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   -300
      Width           =   1275
      Begin VB.CommandButton Command15 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   45
         Picture         =   "NUEMPRE09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Genera Reporte Solicitado"
         Top             =   390
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
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
      Height          =   315
      Left            =   6390
      TabIndex        =   4
      Top             =   585
      Width           =   180
   End
   Begin VB.TextBox Text1 
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
      Left            =   945
      TabIndex        =   1
      Top             =   195
      Width           =   4005
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "NUEMPRE09.frx":030A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "NUEMPRE09.frx":053E
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   945
      TabIndex        =   3
      Top             =   585
      Width           =   5445
   End
   Begin VB.Label Label2 
      Caption         =   "Ruta:"
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
      Left            =   390
      TabIndex        =   2
      Top             =   630
      Width           =   1380
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Left            =   90
      TabIndex        =   0
      Top             =   240
      Width           =   1050
   End
End
Attribute VB_Name = "NUEMPRE09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  '
  CD1.ShowOpen
  '
  
  If CD1.FileName <> "" Then
     ARCH% = Len(CD1.FileTitle)
     Total% = Len(CD1.FileName)
     Ruta% = Total% - ARCH% - 1
     '
     CADENA$ = Mid(CD1.FileName, 1, Ruta%)
     Label3 = CADENA$
  End If
  '
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   '
   If TRIM$(Text1) = "" Then
      Call COMUNI("Falta Ingresar Empresa")
      GoTo 99
   End If
   '
   If TRIM$(Label3) = "" Then
      Call COMUNI("Falta Ingresar Ruta")
      GoTo 99
   End If
   '
   NX% = FILEOPEN(App.Path + "\FLEXMUL.DRV", 0, 79)
   Dim FMY As String * 79
   '
   RASO$ = AJUSTI$(Left$(Text1, 30), 30)
   PACCE$ = AJUSTI$(Mid(Label3, 3), 46)
   Z1$ = String$(79, 0)
   Call REPLA(Z1$, ENCRIPTA$(RASO$), 1, 30)
   Call REPLA(Z1$, ENCRIPTA$(PACCE$), 32, 46)
   SUCAR% = 0
   For J% = 1 To 77
     SUCAR% = SUCAR% + Asc(Mid$(Z1$, J%, 1))
   Next J%
   Call REPLA(Z1$, MKI$(SUCAR%), 78, 2)
   V& = (LOF(NX%) / 79) + 1
   LSet FMY$ = Z1$
   Put #NX%, V&, FMY$
   '
   On Error Resume Next
   Call INICONTA07.CARLISEMP
   If Err Then
      Call INICONTA12.CARLISEMP
   End If
   On Error GoTo 0
   CONN_SQLS.Label1 = Label3
   CONN_SQLS.TxtBasedeDatos = TRIM$(RASO$)
   '
   LUDAT$ = Label3 + "\"
   CONN_SQLS.Show 1
   Call COMUNI("Re-Inicie Ahora la Aplicacion")
   Call FINAL("")
   '
99
   Command15.Enabled = True
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Form_Load()
   Text1 = ""
   Label3 = ""
   Call APLICACIONSKINS(Me, Picture7)
End Sub
