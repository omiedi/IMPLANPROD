VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form SELPEDIPEN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Ordenes de Venta"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7290
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   12
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
      Height          =   2580
      Left            =   80
      Style           =   1  'Checkbox
      TabIndex        =   0
      ToolTipText     =   "Tilde los Pedidos a Considerar"
      Top             =   900
      Width           =   6225
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   80
      ScaleHeight     =   405
      ScaleWidth      =   6195
      TabIndex        =   6
      Top             =   495
      Width           =   6225
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1215
         TabIndex        =   17
         Top             =   105
         Width           =   1215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   250
         TabIndex        =   16
         Top             =   110
         Width           =   975
      End
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
         ForeColor       =   &H00404000&
         Height          =   225
         Left            =   220
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   10515
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   6400
      ScaleHeight     =   5745
      ScaleWidth      =   1485
      TabIndex        =   3
      Top             =   0
      Width           =   1515
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H002B9973&
         Height          =   200
         Left            =   120
         ScaleHeight     =   135
         ScaleWidth      =   585
         TabIndex        =   14
         Top             =   3150
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   15
            Top             =   0
            Width           =   12000
         End
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
         Picture         =   "SELPEDIPEN.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   1155
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
         Picture         =   "SELPEDIPEN.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Aprueba la Selección"
         Top             =   1800
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
         Picture         =   "SELPEDIPEN.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   240
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
      OleObjectBlob   =   "SELPEDIPEN.frx":075E
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   4800
      OleObjectBlob   =   "SELPEDIPEN.frx":0992
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label5 
      Height          =   435
      Left            =   4200
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label2 
      Height          =   435
      Left            =   4830
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Marque (tilde) los pedidos que desea incluir."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   80
      TabIndex        =   9
      Top             =   150
      Width           =   6345
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
      Top             =   1140
      Width           =   1695
   End
End
Attribute VB_Name = "SELPEDIPEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTEXIT_Click()
   Label5 = "0"
   LICOPEN.Clear
   Unload Me
End Sub
'

Sub BOTREC_Click()
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
   '
   Hide
End Sub

Private Sub Command1_Click()
   command1.Enabled = False
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
   command1.Enabled = True
End Sub

'LA RUTINA SE MOVIO AL FORFASQL
'Sub ASIGNALOTES()
'   'SELECCION DE LOTES A ASIGNAR AL PEDIDO
'   REGIS% = 0
'   ASIGLOT07.List1.Clear
'   ASIGLOT07.List2.Clear
'   For U& = 1 To ULTREG&("!CARDEFAC")
'       XX$ = REGLEIDO$("!CARDEFAC", U&)
'       CCII% = CVI(Mid$(XX$, 5, 2))
'       If CCII% > 0 And ESTRAZABLE%(CCII%) = 1 Then
'         TXTX$ = ""
'         Call REPLA(TXTX$, CODEXT$(CCII%), 1, 16)
'         Call REPLA(TXTX$, DESCRIT$(CCII%), 18, 33)
'         SALPEND = CVS(Mid$(XX$, 1, 4))
'         Call REPLA(TXTX$, FORMATNUM$(SALPEND, "F10.1"), 54, 12)
'         Call REPLA(TXTX$, TRIM$(FORMATNUM$(U&, "I5")), 80, 5)
'         ASIGLOT07.List1.AddItem (TXTX$)
'         REGIS% = REGIS% + 1
'       End If
'   Next U&
'   If REGIS% > 0 Then
'      ASIGLOT07.Show 1
'   End If
'   '
'End Sub

Private Sub Form_Load()
    If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
'      Picture16.Visible = True
'      Picture16.ZOrder 0
'      Me.BackColor = &HFCD7B6
'      MARCOBARRA.Top = 3250
'
'      Set MARCOBARRA.Container = Me.Picture16
'      MARCOBARRA.ZOrder 0
    End If

   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me, Picture2)
End Sub

Private Sub grabar_Click()
   Call BOTREC_Click
End Sub

Private Sub revertirseleccion_Click()
    Call Command1_Click
End Sub

Private Sub salir_Click()
   Call BOTEXIT_Click
End Sub
