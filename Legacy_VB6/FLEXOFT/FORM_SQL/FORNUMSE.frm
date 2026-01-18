VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FORNUMSE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EEEEDD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10320
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   6480
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   17
      Top             =   6600
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EEEEDD&
      BorderStyle     =   0  'None
      Height          =   100
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   105
      ScaleWidth      =   15000
      TabIndex        =   16
      Top             =   7875
      Width           =   15000
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
      Height          =   300
      Left            =   4725
      TabIndex        =   15
      Top             =   4830
      Width           =   175
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FFFFC0&
      Height          =   1095
      Left            =   3255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   5880
      Width           =   1485
   End
   Begin VB.TextBox Text2 
      Height          =   435
      Left            =   840
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   5460
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   105
      Top             =   4200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controles de Prueba no Visibles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   5460
      TabIndex        =   11
      Top             =   4515
      Visible         =   0   'False
      Width           =   4320
      Begin VB.ListBox LIPRUEBA 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   210
         TabIndex        =   12
         Top             =   315
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.Label TEPRUANCH 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "TEPRUANCH"
         Height          =   225
         Left            =   210
         TabIndex        =   13
         Top             =   1050
         Visible         =   0   'False
         Width           =   1080
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   9490
      ScaleHeight     =   75
      ScaleWidth      =   675
      TabIndex        =   9
      Top             =   120
      Width           =   735
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   3255
      TabIndex        =   4
      Top             =   4800
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   4110
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7250
      _Version        =   393216
      Rows            =   50
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   8433840
      BackColorSel    =   -2147483643
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame FRAMEBOT1 
      BackColor       =   &H00EEEEDD&
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1520
      Left            =   9480
      TabIndex        =   0
      Top             =   360
      Width           =   750
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
         Height          =   525
         Left            =   90
         Picture         =   "FORNUMSE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Graba - Guardar Modificaciones"
         Top             =   820
         Width           =   540
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
         Height          =   525
         Left            =   90
         Picture         =   "FORNUMSE.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Cancela los Datos Ingresados"
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   8280
      Stretch         =   -1  'True
      ToolTipText     =   "Doble-Click para Bloquear Posición"
      Top             =   6480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label TEPRUTEX 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   225
      Left            =   3255
      TabIndex        =   8
      Top             =   5355
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   225
      Left            =   2415
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label PROFILE 
      Caption         =   "PROFILE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   315
      TabIndex        =   2
      Top             =   5040
      Visible         =   0   'False
      Width           =   1380
   End
End
Attribute VB_Name = "FORNUMSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RESPUESTA As String

Private Sub BOTEXIT_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub BOTREC_Click()
    GRID.TEXT = Text1.TEXT
    '
    For II& = 1 To ULTREG("!CARDEFAC")
        YY$ = REGLEIDO$("!CARDEFAC", II&)
        GRID.Row = II&: GRID.COL = 2
        NSERIE$ = TRIM$(GRID.TEXT)
        Call REPLA(YY$, NSERIE$, 139, 64)
        Call GRAREG("!CARDEFAC", YY$, II&)
    Next II&
    Call CIERRARCH("*.*")
    '
    RESPUESTA = "OK"
    Hide
End Sub

Private Sub Form_Activate()
    Text1.SetFocus
End Sub

Private Sub Form_Load()
    RESPUESTA = "NOOK"
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
End Sub

Sub GRID_Click()

10  If GRID.COL = 2 Then
        GRID.COL = 0
        If GRID.Row <= ULTREG&("!CARDEFAC") Then
            GRID.COL = 2
            Text1.Top = GRID.CellTop + GRID.Top
            Text1.Left = GRID.CellLeft + GRID.Left
            Text1.Height = GRID.CellHeight - 10
            Text1.Width = GRID.CellWidth
            Text1.TEXT = TRIM$(GRID.TEXT)
            Text1.Visible = True
            On Error Resume Next
            Command1.SetFocus
            Text1.SetFocus
            On Error GoTo 0
        Else
            GRID.Row = 1
            GRID.COL = 2
            GoTo 10
        End If
    End If
    
End Sub

Private Sub GRID_GotFocus()
    Text1.SetFocus
End Sub

Private Sub GRID_KeyPress(KeyAscii As Integer)
A = B
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        GRID.TEXT = TRIM$(Text1.TEXT)
        GRID.Refresh
        GRID.Row = GRID.Row + 1
        Call GRID_Click
    End If
    If KeyCode = 38 Then
        GRID.TEXT = TRIM$(Text1.TEXT)
        GRID.Refresh
        GRID.Row = GRID.Row - 1
        Call GRID_Click
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        GRID.TEXT = TRIM$(Text1.TEXT)
        GRID.Refresh
        GRID.Row = GRID.Row + 1
        Call GRID_Click
    End If

End Sub
