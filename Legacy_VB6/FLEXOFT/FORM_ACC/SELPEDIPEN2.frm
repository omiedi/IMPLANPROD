VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form SELPEDIPEN2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Trabajos a Facturar"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   12720
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   12720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   14
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ListBox LICOPEN 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      ToolTipText     =   "Tilde los Pedidos a Considerar"
      Top             =   960
      Width           =   11385
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   11325
      TabIndex        =   8
      Top             =   495
      Width           =   11385
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
         TabIndex        =   9
         Top             =   105
         Width           =   10515
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   5775
      Left            =   11640
      ScaleHeight     =   5715
      ScaleWidth      =   1095
      TabIndex        =   3
      Top             =   0
      Width           =   1155
      Begin VB.PictureBox Picture16 
         Height          =   5100
         Left            =   0
         ScaleHeight     =   5040
         ScaleWidth      =   1140
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
         Begin VB.Image Image3 
            Height          =   5205
            Left            =   0
            Picture         =   "SELPEDIPEN2.frx":0000
            Stretch         =   -1  'True
            Top             =   -120
            Width           =   1095
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
         Left            =   240
         Picture         =   "SELPEDIPEN2.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   915
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   200
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   5
         Top             =   3570
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   12000
         End
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
         Left            =   240
         Picture         =   "SELPEDIPEN2.frx":24E2
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Aprueba la Selección"
         Top             =   2880
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
         Left            =   240
         Picture         =   "SELPEDIPEN2.frx":27EC
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -180
         Picture         =   "SELPEDIPEN2.frx":2AF6
         Top             =   3915
         Width           =   1515
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
      OleObjectBlob   =   "SELPEDIPEN2.frx":4A18
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "SELPEDIPEN2.frx":4C4C
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label5 
      Height          =   435
      Left            =   4200
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label2 
      Height          =   435
      Left            =   4830
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Marque (tilde) los Items que desea incluir: "
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
      Left            =   105
      TabIndex        =   11
      Top             =   105
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
      TabIndex        =   10
      Top             =   1140
      Width           =   1695
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu grabar 
         Caption         =   "Aceptar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu revertirseleccion 
         Caption         =   "Revertir Selección"
      End
      Begin VB.Menu ss 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "SELPEDIPEN2"
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

Private Sub Command3_Click()

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
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      MARCOBARRA.Top = 3250
      
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
    End If

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
