VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FORSELBRE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C7D9CC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recepciones a Facturar"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   8085
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H0050765A&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   7200
      ScaleHeight     =   4545
      ScaleWidth      =   1065
      TabIndex        =   8
      Top             =   0
      Width           =   1095
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H000C93DC&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   13
         Top             =   3240
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0076C9F5&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   14
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton BOTEXIT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "FORSELBRE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   120
         Width           =   650
      End
      Begin VB.CommandButton BOTREC 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "FORSELBRE.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Aprueba la Pre-Selección"
         Top             =   800
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "FORSELBRE.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   1480
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "FORSELBRE.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Ver y Re-Imprimir"
         Top             =   2160
         Width           =   650
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   105
      ScaleHeight     =   555
      ScaleWidth      =   6885
      TabIndex        =   2
      Top             =   3720
      Width           =   6945
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
         Height          =   285
         Left            =   5265
         TabIndex        =   4
         Top             =   150
         Width           =   195
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   4320
         TabIndex        =   5
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Límite de Programación hasta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   180
         Width           =   3375
      End
   End
   Begin VB.ListBox LIMATSEL 
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
      Height          =   3600
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   50
      Width           =   6945
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   3045
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2205
      Visible         =   0   'False
      Width           =   1065
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "FORSELBRE.frx":0A68
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "FORSELBRE.frx":0AE2
      Top             =   1200
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7050
      Y1              =   4340
      Y2              =   4340
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Aprobar 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu S 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu editar 
      Caption         =   "Editar"
      Begin VB.Menu revertirseelccion 
         Caption         =   "Revertir Selección"
      End
   End
   Begin VB.Menu VER 
      Caption         =   "Ver"
      Begin VB.Menu VERIMAGEN 
         Caption         =   "Ver y Re-Imprimir Recepción "
      End
   End
End
Attribute VB_Name = "FORSELBRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Aprobar_Click()
    Call BOTREC_Click
End Sub

Private Sub BOTEXIT_Click()
    LIMATSEL.Clear
    Label2.Caption = ""
    Hide
End Sub
'

Private Sub BOTREC_Click()
   Hide
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   TPI = LIMATSEL.TopIndex
   For U& = 1 To LIMATSEL.ListCount
     If LIMATSEL.Selected(U& - 1) = True Then
         LIMATSEL.Selected(U& - 1) = False
       Else
         LIMATSEL.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   LIMATSEL.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   Command1.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   If LIMATSEL.ListIndex >= 0 Then
     Y$ = LIMATSEL.LIST(LIMATSEL.ListIndex)
     NBORE& = Val(Mid$(Y$, 4, 6))
     Call SHOWBORE(NBORE&)
     GoTo 99
   End If
3  Call SELECHO("INDIBOL")
   NPX& = Val(VALACT1$("INDIBOL"))
   If NPX& > 0 Then
     Call SHOWBORE(NPX&)
     GoTo 3
   End If
   '
99 Command4.Enabled = True
End Sub

Private Sub Command9_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Label2 = VALACT1$("SELFECHA")
   End If
End Sub
'
Sub SHOWBORE(NPX&)
    '
    TIDOCUM$ = "Boleta de Recepción"
    NUDOCU$ = TRIM$(Str$(NPX&))
    NUDOCU$ = " " + NUDOCU$ + " -"
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        If InStr(XX$, NUDOCU$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    Exit Sub
    '
5   Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
    '
End Sub

Private Sub Form_Load()
'    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
'      Me.Picture16.Visible = True
'      Me.BackColor = &HFCD7B6
'      MARCOBARRA.Top = 3300
'      Set MARCOBARRA.Container = Me.Picture16
'      MARCOBARRA.ZOrder 0
'    End If
'
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
End Sub

Private Sub revertirseelccion_Click()
   Call Command1_Click
End Sub

Private Sub salir_Click()
     Call BOTEXIT_Click
End Sub

Private Sub VERIMAGEN_Click()
   Call Command4_Click
End Sub
