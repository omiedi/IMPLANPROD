VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FRMCOLOR 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Colores"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6045
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   7
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00800000&
      Height          =   3900
      Left            =   5160
      ScaleHeight     =   3840
      ScaleWidth      =   1245
      TabIndex        =   1
      Top             =   -120
      Width           =   1305
      Begin VB.CommandButton Command2 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Left            =   90
         Picture         =   "FRMCOLOR.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Búsqueda por  Texto"
         Top             =   960
         Width           =   650
      End
      Begin VB.CommandButton BOTEXIT 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Left            =   90
         Picture         =   "FRMCOLOR.frx":1366
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cancelar Selección"
         Top             =   225
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
         Height          =   680
         Left            =   90
         Picture         =   "FRMCOLOR.frx":1670
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Aprobar Selección"
         Top             =   2085
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   110
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   2
         Top             =   2790
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   12000
         End
      End
   End
   Begin VB.ListBox LICHESEL 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2865
      Left            =   50
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   90
      Width           =   5000
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FRMCOLOR.frx":197A
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "FRMCOLOR.frx":1BAE
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "FRMCOLOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public COLOR$
Private Sub BOTEXIT_Click()
    Hide
End Sub
'
Private Sub BOTREC_Click()

    COLOR$ = ""
    VUELTA% = 0
    For i& = 1 To LICHESEL.ListCount
       If LICHESEL.Selected(i& - 1) = True Then
        If VUELTA% = 0 Then
            COLOR$ = COLOR$ + TRIM$(Mid$(LICHESEL.List(i& - 1), 7, 24))
            VUELTA% = 1
        Else
            COLOR$ = COLOR$ & "; " & TRIM$(Mid$(LICHESEL.List(i& - 1), 7, 24))
       End If
      End If
    Next i&
    '
    Call BOTEXIT_Click
End Sub
'
Private Sub Command2_Click()
   Static ABUSQUE$
   ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", "Introduzca Argumento de Búsqueda", ABUSQUE$)))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     INI& = LICHESEL.ListIndex
10   For U& = INI& + 2 To LICHESEL.ListCount
       If InStr(UCase$(LICHESEL.List(U& - 1)), ABUSQUE$) > 0 Then
         LICHESEL.ListIndex = U& - 1
         Exit Sub
       End If
     Next U&
     If INI& >= 0 Then
       INI& = -1
       GoTo 10
     End If
     '
     Call MENSERR(24, "Argumento no Encontrado.")
   End If
End Sub

Private Sub Form_Load()
   Call APLICACIONSKINS(Me, Picture2)
   SENDBYMAIL$ = ""
End Sub
