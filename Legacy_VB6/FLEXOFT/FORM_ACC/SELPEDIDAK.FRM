VERSION 5.00
Begin VB.Form SELPEDIDAK 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Ordenes de Venta"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10020
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4095
      TabIndex        =   0
      ToolTipText     =   "<Enter> en Vacío Para Incluir Todos"
      Top             =   60
      Width           =   4740
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
      Height          =   2610
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   1
      ToolTipText     =   "Tilde los Pedidos a Considerar"
      Top             =   900
      Width           =   8745
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   8685
      TabIndex        =   9
      Top             =   495
      Width           =   8745
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "   Nro.  F.Emis   F.Entreg  Descuento        Suc    O/Compra     "
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
         TabIndex        =   10
         Top             =   105
         Width           =   10515
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   5775
      Left            =   9030
      ScaleHeight     =   5715
      ScaleWidth      =   1455
      TabIndex        =   4
      Top             =   0
      Width           =   1515
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
         Picture         =   "SELPEDIDAK.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   1155
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   6
         Top             =   2730
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   7
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
         Left            =   120
         Picture         =   "SELPEDIDAK.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Aprueba la Selección"
         Top             =   1860
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
         Picture         =   "SELPEDIDAK.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "SELPEDIDAK.frx":075E
         Top             =   3075
         Width           =   1515
      End
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   10200
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   5400
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label5 
      Height          =   435
      Left            =   4200
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label2 
      Height          =   435
      Left            =   4830
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Marque (tilde) los pedidos que desea incluir: "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   105
      TabIndex        =   12
      Top             =   105
      Width           =   6945
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   8860
      Y1              =   3510
      Y2              =   3510
   End
   Begin VB.Line Line2 
      X1              =   8860
      X2              =   8860
      Y1              =   525
      Y2              =   3465
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
      TabIndex        =   11
      Top             =   1140
      Width           =   1695
   End
End
Attribute VB_Name = "SELPEDIDAK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PORCLICK%
Dim VIENEDETEXT1%
'
Private Sub BOTEXIT_Click()
   Label5 = "0"
   LICOPEN.Clear
   Unload Me
End Sub
'

Private Sub BOTREC_Click()
   '
   Label5 = "0"
   DESCUSE$ = "": SUCUSE$ = ""
   HUBOSUPRI% = 0
   For I = 1 To LICOPEN.ListCount
     If LICOPEN.Selected(I - 1) = True Then
       TTXX$ = LICOPEN.List(I - 1)
       DESCUAC$ = TRIM$(Mid$(TTXX$, 29, 14))
       SUCUAC$ = TRIM$(Mid$(TTXX$, 45, 4))
       If DESCUSE = "" Then DESCUSE$ = DESCUAC$
       If SUCUSE$ = "" Then SUCUSE$ = SUCUAC$
       If DESCUAC$ <> "" Then
          If DESCUAC$ <> DESCUSE$ Then
             LICOPEN.Selected(I - 1) = False
             HUBOSUPRI% = 1
             GoTo 49
          End If
       End If
       '
       If SUCUAC$ <> "" Then
          If SUCUAC$ <> SUCUSE$ Then
             LICOPEN.Selected(I - 1) = False
             HUBOSUPRI% = 1
             GoTo 49
          End If
       End If
       '
       CANSELE% = CANSELE% + 1
     End If
49 Next I
   If HUBOSUPRI% > 0 Then
      Call MENSERR(24, "No Válido Combinar Pedidos\de Distinto Descuento o Sucursal.\   \Revise la Selección.")
      Text1.SetFocus
      Exit Sub
   End If
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

Private Sub Form_Activate()
    Text1.SetFocus
End Sub

Private Sub LICOPEN_Click()
   If VIENEDETEXT1% > 0 Then Exit Sub
   TTXX$ = ""
   For KKI% = 1 To LICOPEN.ListCount
      If LICOPEN.Selected(KKI% - 1) = True Then
        If TTXX$ <> "" Then TTXX$ = TTXX$ + ","
        TTXX$ = TTXX$ + TRIM$(Left$(LICOPEN.List(KKI% - 1), 6))
      End If
   Next KKI%
   '
   Text1 = TTXX$
End Sub

Private Sub Text1_Change()
   If PORCLICK% > 0 Then Exit Sub
   VIENEDETEXT1% = 1
   For KKI% = 1 To LICOPEN.ListCount
      LICOPEN.Selected(KKI% - 1) = False
   Next KKI%
   '
   TTX2$ = Text1.TEXT
   TTX2$ = REPLACAR$(TTX2$, ".", ",")
   TTX2$ = REPLACAR$(TTX2$, "+", ",")
   TTX2$ = REPLACAR$(TTX2$, "/", ",")
   TTX2$ = REPLACAR$(TTX2$, " ", ",")
   '
10 PPXX% = InStr(TTX2$, ",")
   If PPXX% < 1 Then PPXX% = 1 + Len(TTX2$)
   NPX = XVALO(Left$(TTX2$, PPXX% - 1))
   TTX2$ = Mid$(TTX2$, PPXX% + 1)
   For KKI% = 1 To LICOPEN.ListCount
      If XVALO(Left$(LICOPEN.List(KKI% - 1), 6)) = XVALO(NPX) Then
         LICOPEN.Selected(KKI% - 1) = True
      End If
   Next KKI%
   If TTX2$ <> "" Then GoTo 10
   VIENEDETEXT1% = 0
   '
End Sub

Private Sub Text1_GotFocus()
   PORCLICK% = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      If Text1 = "" Then
         VIENEDETEXT1% = 1
         For KKI% = 1 To LICOPEN.ListCount
            LICOPEN.Selected(KKI% - 1) = True
         Next KKI%
         VIENEDETEXT1% = 0
      End If
      '
      BOTREC.SetFocus
      Call BOTREC_Click
   End If
   '
   If KeyAscii% = 27 Then
      If Text1 <> "" Then
         VIENEDETEXT1 = 1
         Text1 = ""
         Exit Sub
      End If
   End If
   '
End Sub

Private Sub Text1_LostFocus()
   PORCLICK% = 1
End Sub
