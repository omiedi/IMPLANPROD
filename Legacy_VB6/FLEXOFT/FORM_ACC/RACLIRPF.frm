VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form RACLIRPF 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remitos pendientes de Facturar"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   4845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   600
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   12
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00624E28&
      Height          =   2850
      Left            =   3960
      ScaleHeight     =   2790
      ScaleWidth      =   1020
      TabIndex        =   9
      Top             =   -105
      Width           =   1080
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   14
         Top             =   1800
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
         Height          =   690
         Left            =   105
         Picture         =   "RACLIRPF.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   240
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Left            =   105
         Picture         =   "RACLIRPF.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   930
         Width           =   650
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E2D3C0&
      Caption         =   "RANGO DE CUENTAS DE CLIENTES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3765
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   195
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1630
         Width           =   3330
      End
      Begin VB.CommandButton Command8 
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
         Left            =   3345
         TabIndex        =   5
         Top             =   1260
         Width           =   175
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2505
         TabIndex        =   4
         Text            =   " "
         Top             =   1260
         Width           =   855
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   195
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   800
         Width           =   3330
      End
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
         Height          =   300
         Left            =   3345
         TabIndex        =   2
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2505
         TabIndex        =   1
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
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
         Left            =   90
         TabIndex        =   8
         Top             =   1365
         Width           =   2325
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
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
         Left            =   120
         TabIndex        =   7
         Top             =   510
         Width           =   2325
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "RACLIRPF.frx":0454
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "RACLIRPF.frx":0688
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu aceptar 
         Caption         =   "Aceptar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "RACLIRPF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aceptar_Click()
   Call Command2_Click
End Sub

Private Sub Command1_Click()
    '
    Unload Me
    '
End Sub

Private Sub Command2_Click()
    '
    RACLIRPF.Hide
    '
End Sub

Private Sub Command8_Click()
    Call SELECHO("DEUDOR12")
    NCMAX = XVALO(VALACT1$("DEUDOR12"))
    If NCMAX > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX))
        Text7.TEXT = rasocli$(NCMAX)
    End If
End Sub

Private Sub Command9_Click()
    '
    Call SELECHO("DEUDOR1")
    NCMIN = XVALO(VALACT1$("DEUDOR12"))
    If NCMIN > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN))
        Text9.TEXT = rasocli$(NCMIN)
    End If
    '
End Sub

Private Sub Picture1_Click()
    '
    Call SELECHO("DEUDOR1")
    NCMAX = XVALO(VALACT1$("DEUDOR12"))
    If NCMAX > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX))
        Text7.TEXT = rasocli$(NCMAX)
    End If
    '
End Sub

Private Sub salir_Click()
    Call Command1_Click
End Sub

Private Sub Text10_DblClick()
   '\\\OT-06-0570-RG-27/12/06///
   Call Command9_Click
   '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub Text10_GotFocus()
   Text10.SelStart = 0
   Text10.SelLength = Len(Text10)
End Sub

Private Sub Text8_Change()
    NCMAX = XVALO(Text8.TEXT)
    Text7.TEXT = rasocli$(NCMAX)
End Sub
Private Sub Text8_DblClick()
    Call SELECHO("DEUDOR12")
    NCMAX = XVALO(VALACT1$("DEUDOR12"))
    If NCMAX > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX))
        Text7.TEXT = rasocli$(NCMAX)
    End If
End Sub

Private Sub Text8_GotFocus()
   Text8.SelStart = 0
   Text8.SelLength = Len(Text8)
End Sub

Private Sub Text9_Click()
    Text10.SetFocus
End Sub
Private Sub Text10_Change()
    NCMAX = XVALO(Text10.TEXT)
    Text9.TEXT = rasocli$(NCMAX)
End Sub
