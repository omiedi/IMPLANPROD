VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form GENORF07 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generación de Orden de Fabricación"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   35
      Top             =   720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   8400
      Picture         =   "GENORF07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "O.K."
      Top             =   3675
      Width           =   750
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Materiales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   315
      TabIndex        =   27
      Top             =   2520
      Width           =   7995
      Begin VB.Frame Frame6 
         Caption         =   "Frame6"
         Height          =   1280
         Left            =   5775
         TabIndex        =   34
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   1280
         Left            =   4830
         TabIndex        =   33
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   1280
         Left            =   3885
         TabIndex        =   32
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1280
         Left            =   1995
         TabIndex        =   31
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame11 
         Caption         =   "Frame3"
         Height          =   1280
         Left            =   6720
         TabIndex        =   30
         Top             =   315
         Width           =   15
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
         Height          =   780
         ItemData        =   "GENORF07.frx":030A
         Left            =   210
         List            =   "GENORF07.frx":030C
         TabIndex        =   28
         Top             =   630
         Width           =   7575
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"GENORF07.frx":030E
         Height          =   330
         Left            =   210
         TabIndex        =   29
         Top             =   315
         Width           =   7575
      End
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7245
      TabIndex        =   25
      Text            =   " "
      Top             =   105
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   315
      TabIndex        =   8
      Top             =   630
      Width           =   7995
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4515
         TabIndex        =   24
         Text            =   " "
         Top             =   420
         Width           =   1065
      End
      Begin VB.CommandButton Command6 
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
         Left            =   5565
         TabIndex        =   22
         Top             =   1260
         Width           =   175
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4515
         TabIndex        =   21
         Top             =   1260
         Width           =   1050
      End
      Begin VB.CommandButton Command7 
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
         Left            =   1890
         TabIndex        =   19
         Top             =   1260
         Width           =   175
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   18
         Top             =   1260
         Width           =   1050
      End
      Begin VB.CommandButton Command5 
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
         Left            =   7665
         TabIndex        =   15
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6615
         TabIndex        =   13
         Text            =   " "
         Top             =   420
         Width           =   1065
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3150
         TabIndex        =   12
         Text            =   " "
         Top             =   420
         Width           =   435
      End
      Begin VB.CommandButton Command4 
         Caption         =   "."
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
         Left            =   2835
         TabIndex        =   11
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   10
         Top             =   420
         Width           =   2010
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3870
         TabIndex        =   23
         Top             =   525
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Entrega Solic.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3150
         TabIndex        =   20
         Top             =   1365
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lanz.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -525
         TabIndex        =   17
         Top             =   1365
         Width           =   1335
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
         Height          =   285
         Left            =   840
         TabIndex        =   16
         Top             =   840
         Width           =   7020
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5970
         TabIndex        =   14
         Top             =   525
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -525
         TabIndex        =   9
         Top             =   525
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00004080&
      Height          =   8460
      Left            =   9225
      ScaleHeight     =   8400
      ScaleWidth      =   1005
      TabIndex        =   3
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command16 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "GENORF07.frx":0397
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Listado de Informe de Produccion por Fecha"
         Top             =   1575
         Width           =   650
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "GENORF07.frx":06A1
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   2625
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   4
         Top             =   3570
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   5
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
         Height          =   640
         Left            =   105
         Picture         =   "GENORF07.frx":0AE3
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   200
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "GENORF07.frx":0C2D
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   845
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "GENORF07.frx":0F37
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   3885
         Width           =   1515
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "GENORF07.frx":2E59
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "GENORF07.frx":308D
      TabIndex        =   36
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. O/F.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5880
      TabIndex        =   26
      Top             =   210
      Width           =   1335
   End
End
Attribute VB_Name = "GENORF07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Command29_Click()
  '
  Call OPNOIN("Ayuda FLEXOFT en Línea")
  '
End Sub

Private Sub Command4_Click()
  '
  Command4.Enabled = False
  Call SELECHO("MASTER")
  CODD$ = VALACT1$("MASTER")
  If CODD$ <> "" Then
     Text1.TEXT = CODD$
     CI% = CODINT%(CODD$)
     Label3.Caption = DESCRIT$(CI%)
  End If
  Command4.Enabled = True
  '
End Sub

Private Sub Command5_Click()
  '
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
     Text4.TEXT = VDEV$
  End If
  '
End Sub

Private Sub Command6_Click()
  '
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
     Text6.TEXT = VDEV$
  End If
   '
End Sub

Private Sub Command7_Click()
  '
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
     Text5.TEXT = VDEV$
  End If
  '
End Sub

Sub Form_Load()
  '
  Call APLICACIONSKINS(Me, Picture14)
  Screen.MousePointer = 11
  Call CENTRAFORM(Me)
  HOII% = HOY(HOS$)
  Text4.TEXT = FECHATEX$(HOII%)
  Screen.MousePointer = 1
  '
End Sub

Private Sub Text1_DblClick()
  '
  Call SELECHO("MASTER")
  CODD$ = VALACT1$("MASTER")
  If CODD$ <> "" Then
     Text1.TEXT = CODD$
     CI% = CODINT%(CODD$)
     Label3.Caption = DESCRIT$(CI%)
  End If
  '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  '
  If KeyAscii = 13 Then
     Text4.SetFocus
  End If
  '
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
  '
  If KeyAscii = 13 Then
     Text5.SetFocus
  End If
  '
End Sub

Private Sub Text3_LostFocus()
  '
  NOPX$ = REPLACAR$(TRIM$(Text3), "-", "/")
  NOPX$ = REPLACAR$(NOPX$, ".", "/")
  Text3 = NOPX$
  '
End Sub

Private Sub Text4_DblClick()
  '
  Call Command5_Click
  '
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
  '
  If KeyAscii% = 13 Then
     Text1.SetFocus
  End If
  '
End Sub

Private Sub Text5_DblClick()
  '
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
     Text5.TEXT = VDEV$
  End If
  '
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
  '
  If KeyAscii = 13 Then
     Text6.SetFocus
  End If
  '
End Sub

Private Sub Text6_DblClick()
  '
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
     Text6.TEXT = VDEV$
  End If
  '
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
  '
  If KeyAscii = 13 Then
     Text7.SetFocus
  End If
  '
End Sub


