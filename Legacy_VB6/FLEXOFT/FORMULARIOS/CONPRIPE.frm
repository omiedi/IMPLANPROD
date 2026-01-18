VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CONPRIPE 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Impresión"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2760
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   875
      Left            =   5760
      Picture         =   "CONPRIPE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Aprueba, Graba e Imprime Pedido"
      Top             =   600
      Width           =   540
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00E2D3C0&
      Caption         =   "IMPRESIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   105
      TabIndex        =   0
      Top             =   480
      Width           =   5595
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Replica por Mail"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   4300
         TabIndex        =   6
         Top             =   315
         Width           =   1170
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Impresión Fantasma"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   2950
         TabIndex        =   3
         Top             =   315
         Width           =   1170
      End
      Begin VB.CheckBox Check8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Imprime O/Despacho"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   1400
         TabIndex        =   2
         Top             =   315
         Width           =   1560
      End
      Begin VB.CheckBox Check7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Imprime O/Venta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   240
         TabIndex        =   1
         Top             =   315
         Value           =   1  'Checked
         Width           =   1140
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "CONPRIPE.frx":030A
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   4800
      OleObjectBlob   =   "CONPRIPE.frx":053E
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Pto.Orden de Venta: : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Doble Click Para Configurar Puerto de Impr. - Orden de Despacho"
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pto. Orden de Despacho: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Doble Click Para Configurar Puerto de Impr. - Orden de Despacho"
      Top             =   240
      Width           =   6375
   End
   Begin VB.Label Label1 
      Height          =   435
      Left            =   105
      TabIndex        =   4
      Top             =   1890
      Width           =   1380
   End
End
Attribute VB_Name = "CONPRIPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command14_Click()
    Label1.Caption = "OK"
    Hide
End Sub

Private Sub Form_Load()
    '
    If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
      Me.Frame10.BackColor = &HFCD7B6
      Me.Check1.BackColor = &HFCD7B6
      Me.Check2.BackColor = &HFCD7B6
      Me.Check7.BackColor = &HFCD7B6
      Me.Check8.BackColor = &HFCD7B6
      '
    End If

    Check7.Value = 1
    If Control$("PEDIDO", "PRINOVEN") = "NO" Then Check7.Value = 0
    Check8.Value = 0
    If Control$("PEDIDO", "PRINODES") = "SI" Then Check8.Value = 1
    Check2.Value = 0
    If Control$("PEDCLIEN", "SUPRINT") = "SI" Then Check2.Value = 1
    '
    PDES$ = Control$(IDENTER$, "PUERDESP")
    If PDES$ = "" Then PDES$ = Control$("*", "PUERDESP")
    Label2.Caption = "Pto. Orden de Despacho: " & PDES$
    POVE$ = Control$(IDENTER$, "PUERORVE")
    If POVE$ = "" Then POVE$ = Control$("*", "PUERORVE")
    Label3.Caption = "Pto. Orden de Venta: " & POVE$

    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)

End Sub

Private Sub Label2_DblClick()
    IMPREDESPA.Show 1
End Sub
Private Sub Label3_DblClick()
    IMPREDESPA.Show 1
End Sub

