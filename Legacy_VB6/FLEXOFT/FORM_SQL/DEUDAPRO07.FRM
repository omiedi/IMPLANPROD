VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form DEUDAPRO07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C7D9CC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Composición de Deuda del Proveedor"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1440
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H0050765A&
      ForeColor       =   &H80000008&
      Height          =   4980
      Left            =   9840
      ScaleHeight     =   4950
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   0
      Width           =   1905
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
         Picture         =   "DEUDAPRO07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Búsqueda por Texto del Detalle"
         Top             =   1480
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
         Height          =   690
         Left            =   100
         Picture         =   "DEUDAPRO07.frx":1366
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Selección / deselección Automática"
         Top             =   800
         Width           =   650
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
         Height          =   690
         Left            =   100
         Picture         =   "DEUDAPRO07.frx":14B0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancelar la Aplicación"
         Top             =   105
         Width           =   650
      End
   End
   Begin VB.ListBox LICHESEL 
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
      Height          =   3090
      ItemData        =   "DEUDAPRO07.frx":17BA
      Left            =   80
      List            =   "DEUDAPRO07.frx":17BC
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   480
      Width           =   9675
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
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2400
      OleObjectBlob   =   "DEUDAPRO07.frx":17BE
      Top             =   3600
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   240
      OleObjectBlob   =   "DEUDAPRO07.frx":19F2
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label87 
      BackStyle       =   0  'Transparent
      Caption         =   "Emisión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      Top             =   230
      Width           =   840
   End
   Begin VB.Label Label86 
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   310
      TabIndex        =   16
      Top             =   230
      Width           =   1455
   End
   Begin VB.Label Label85 
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
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   225
      Width           =   840
   End
   Begin VB.Label Label84 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Importe Original"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   14
      Top             =   225
      Width           =   1440
   End
   Begin VB.Label Label83 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimiento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6890
      TabIndex        =   13
      Top             =   230
      Width           =   1080
   End
   Begin VB.Label Label82 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8970
      TabIndex        =   12
      Top             =   230
      Width           =   840
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   4305
      TabIndex        =   11
      Top             =   3720
      Width           =   2010
   End
   Begin VB.Label LabeL15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Height          =   300
      Left            =   7560
      TabIndex        =   10
      Top             =   3720
      Width           =   2010
   End
   Begin VB.Label Label1 
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "SELECCIONADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3000
      TabIndex        =   9
      Top             =   3840
      Width           =   1305
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL DEUDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6450
      TabIndex        =   8
      Top             =   3840
      Width           =   1020
   End
   Begin VB.Label Label88 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   80
      TabIndex        =   18
      Top             =   190
      Width           =   9675
   End
End
Attribute VB_Name = "DEUDAPRO07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTEXIT_Click()
    LICHESEL.Clear
    label2.Caption = ""
    LabeL15.Caption = ""
    Hide
End Sub
'


Private Sub Command1_Click()
   Command1.Enabled = False
   TPI = LICHESEL.TopIndex
   For U& = 1 To LICHESEL.ListCount
     If LICHESEL.Selected(U& - 1) = True Then
         LICHESEL.Selected(U& - 1) = False
       Else
         LICHESEL.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   LICHESEL.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   Command1.Enabled = True

End Sub


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

Private Sub Form_Activate()
    Call SUCHESEL
    If LICHESEL.ListCount > 0 Then LICHESEL.ListIndex = 0
End Sub

Sub SUCHESEL()
    SUCHESE# = 0
    For i& = 1 To LICHESEL.ListCount
       If LICHESEL.Selected(i& - 1) = True Then
         SUCHESE# = SUCHESE# + XVALO(Mid$(LICHESEL.List(i& - 1), 65, 12))
       End If
    Next i&
    label2.Caption = CSTRING$(MKD$(SUCHESE#), 4, 11, 2, 2)
    label2.Refresh
End Sub


Private Sub Form_Load()
  
  UTILIZA_SKIN% = 1
  Call APLICACIONSKINS(Me)
  
End Sub

Private Sub LICHESEL_Click()
   Call SUCHESEL
End Sub




