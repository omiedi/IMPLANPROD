VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FRMLIST1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0095C5E1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de cuentas contables"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7365
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   4320
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ListBox LIST1 
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
      Height          =   4875
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      ToolTipText     =   "Tilde las Cuentas Contables a Considerar"
      Top             =   240
      Width           =   6250
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H002C3D46&
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   6480
      ScaleHeight     =   5745
      ScaleWidth      =   1485
      TabIndex        =   2
      Top             =   0
      Width           =   1515
      Begin VB.PictureBox MARCOBARRA 
         Height          =   195
         Left            =   100
         ScaleHeight     =   135
         ScaleWidth      =   585
         TabIndex        =   12
         Top             =   5280
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   13
            Top             =   0
            Width           =   12000
         End
      End
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
         Height          =   690
         Left            =   100
         Picture         =   "FRMLIST1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Búsqueda por Número, Banco, Importe, etc."
         Top             =   800
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
         Height          =   690
         Left            =   100
         Picture         =   "FRMLIST1.frx":1366
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   1490
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
         Height          =   690
         Left            =   100
         Picture         =   "FRMLIST1.frx":14B0
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Aprueba la Selección"
         Top             =   4320
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
         Height          =   690
         Left            =   100
         Picture         =   "FRMLIST1.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   120
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FRMLIST1.frx":1AC4
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   2640
      OleObjectBlob   =   "FRMLIST1.frx":1CF8
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   345
      Left            =   5040
      TabIndex        =   10
      Top             =   5280
      Width           =   1320
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad Seleccionada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2640
      TabIndex        =   9
      Top             =   5430
      Width           =   2280
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
      TabIndex        =   5
      Top             =   1140
      Width           =   1695
   End
End
Attribute VB_Name = "FRMLIST1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTEXIT_Click()
   Label5 = "0"
   LIST1.Clear
   Unload Me
End Sub
'

Sub BOTREC_Click()
   '
   Call ACTUSELECCIONADOS
   '
99 Label2 = "OK"
   '
   Hide
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   TPI = LIST1.TopIndex
   For U& = 1 To LIST1.ListCount
     If LIST1.Selected(U& - 1) = True Then
         LIST1.Selected(U& - 1) = False
       Else
         LIST1.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   LIST1.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   Command1.Enabled = True
End Sub



Private Sub Form_Load()
   Call APLICACIONSKINS(Me, Picture2)
End Sub
Private Sub Command2_Click()
   Static ABUSQUE$
   ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", "Introduzca Argumento de Búsqueda", ABUSQUE$)))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     INI& = LIST1.ListIndex
10   For U& = INI& + 2 To LIST1.ListCount
       If InStr(UCase$(LIST1.List(U& - 1)), ABUSQUE$) > 0 Then
         LIST1.ListIndex = U& - 1
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

   Sub ACTUSELECCIONADOS()
    Label5 = "0"
    For i = 1 To LIST1.ListCount
      If LIST1.Selected(i - 1) = True Then
        CANSELE% = CANSELE% + 1
      End If
    Next i
    Label5 = TRIM$(Str$(CANSELE%))
   End Sub

Private Sub LIST1_ItemCheck(Item As Integer)
   Call ACTUSELECCIONADOS
End Sub
