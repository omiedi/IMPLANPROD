VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form SELOCOPEN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aprobación  de Ordenes de Compra"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   9525
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   13
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ListBox LICOPEN 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      ToolTipText     =   "Tilde las Ordenes de Compra a Aprobar"
      Top             =   900
      Width           =   8325
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   120
      ScaleHeight     =   405
      ScaleWidth      =   8295
      TabIndex        =   6
      Top             =   495
      Width           =   8325
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "  Nro.   Referencia     "
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   220
         TabIndex        =   7
         Top             =   105
         Width           =   10515
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   4140
      Left            =   8640
      ScaleHeight     =   4080
      ScaleWidth      =   1455
      TabIndex        =   3
      Top             =   0
      Width           =   1515
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   15
         Top             =   3360
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "SELOCOPEN.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Ver por Pantalla Orden de Compra Activa"
         Top             =   840
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "SELOCOPEN.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   1470
         Width           =   650
      End
      Begin VB.CommandButton BOTREC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "SELOCOPEN.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Aprueba la Selección"
         Top             =   2205
         Width           =   650
      End
      Begin VB.CommandButton BOTEXIT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "SELOCOPEN.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   120
         Width           =   650
      End
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   10200
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   1065
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "SELOCOPEN.frx":0A68
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   195
      Index           =   0
      Left            =   5040
      OleObjectBlob   =   "SELOCOPEN.frx":0C9C
      TabIndex        =   14
      Top             =   240
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label Label5 
      Height          =   435
      Left            =   7350
      TabIndex        =   11
      Top             =   5000
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label2 
      Height          =   435
      Left            =   7980
      TabIndex        =   10
      Top             =   5000
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Marque (tilde) las Ordenes de Compra que Desea Aprobar: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   105
      TabIndex        =   9
      Top             =   105
      Width           =   4665
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   7440
      TabIndex        =   8
      Top             =   5000
      Width           =   1695
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu dfs 
         Caption         =   "Revertir Selección"
      End
      Begin VB.Menu aprobarselecccion 
         Caption         =   "Aprbar Selección"
      End
   End
   Begin VB.Menu ver 
      Caption         =   "Ver"
      Begin VB.Menu vercompraseleccionada 
         Caption         =   "Ver Orden de Compra Seleccionada"
      End
   End
End
Attribute VB_Name = "SELOCOPEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aprobarselecccion_Click()
   Call BOTREC_Click
End Sub

Private Sub BOTEXIT_Click()
   Label5 = "0"
   LICOPEN.Clear
   Unload Me
End Sub
'

Private Sub BOTREC_Click()
   '
   Label5 = "0"
   For i = 1 To LICOPEN.ListCount
     If LICOPEN.Selected(i - 1) = True Then
       CANSELE% = CANSELE% + 1
     End If
   Next i
   Label5 = TRIM$(Str$(CANSELE%))
   '
99 Label2 = "OK"
   Hide
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   TPI = LICOPEN.TopIndex
   For U& = 1 To LICOPEN.ListCount
     If LICOPEN.Selected(U& - 1) = True Then
         LICOPEN.Selected(U& - 1) = False
       Else
         LICOPEN.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   LICOPEN.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   Command1.Enabled = True
End Sub

Private Sub Command4_Click()
    '
    Command4.Enabled = False
    ILISTA& = LICOPEN.ListIndex
    If ILISTA& >= O Then
      LISTACT$ = LICOPEN.TEXT
      NPX& = XVALO(Left$(LISTACT$, 6))
      If NPX& > 0 Then
         '
         TIDOCUM$ = "Nota de Pedido"
         NUDOCU$ = TRIM$(Str$(NPX&))
         While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
         NUDOCU1$ = " " + NUDOCU$ + " "
         NUDOCU2$ = " " + NUDOCU$ + "-"
         Screen.MousePointer = 11
         '
         For UI& = ULTREG&("PDOCLST") To 1 Step -1
           XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
           If InStr(XX$, TIDOCUM$) > 4 Then
             PPXX% = InStr(XX$, "Nro.")
             If PPXX% > 0 Then
               XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
             End If
             If InStr(XX$, NUDOCU1$) > 4 Or InStr(XX$, NUDOCU2$) > 4 Then
               DOCUNU& = CVS(Left$(XX$, 4))
               GoTo 5
             End If
           End If
         Next UI&
         Screen.MousePointer = 1
         Call COMUNI("Documento no Encontrado")
         GoTo 99
         '
5        Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
         '
      End If
    End If
    '
99  Command4.Enabled = True
    '
End Sub

Private Sub dfs_Click()
   Call Command1_Click
End Sub

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture6.Visible = True
      Me.BackColor = &HFCD7B6
      MARCOBARRA.Top = 3600
      'Set MARCOBARRA.Container = Me.Picture6

    End If
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
End Sub


Private Sub vercompraseleccionada_Click()
    Call Command4_Click
End Sub
