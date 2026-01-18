VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form RECUNOBA 
   BackColor       =   &H0095C5E1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recuperación Comprobante No Balanceado"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   16
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0095C5E1&
      Caption         =   "BALANCE AUTOMÁTICO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   105
      TabIndex        =   11
      Top             =   3620
      Width           =   6105
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
         Left            =   2310
         TabIndex        =   12
         Top             =   315
         Width           =   175
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   300
         Left            =   945
         TabIndex        =   15
         Top             =   315
         Width           =   1365
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
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
         Left            =   -1470
         TabIndex        =   14
         Top             =   400
         Width           =   2325
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   300
         Left            =   2600
         TabIndex        =   13
         Top             =   315
         Width           =   3360
      End
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   105
      TabIndex        =   9
      Top             =   5460
      Width           =   9150
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
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
      Left            =   7770
      TabIndex        =   8
      Top             =   4410
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aprobar"
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
      Left            =   6405
      TabIndex        =   7
      Top             =   4410
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   105
      ScaleHeight     =   405
      ScaleWidth      =   9120
      TabIndex        =   1
      Top             =   315
      Width           =   9150
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"RECUNOBA.frx":0000
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   30
         TabIndex        =   2
         Top             =   105
         Width           =   11565
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      Left            =   105
      TabIndex        =   0
      Top             =   735
      Width           =   9150
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1680
      OleObjectBlob   =   "RECUNOBA.frx":009F
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   -360
      OleObjectBlob   =   "RECUNOBA.frx":02D3
      TabIndex        =   17
      Top             =   1560
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   210
      TabIndex        =   10
      Top             =   5040
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"RECUNOBA.frx":0331
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1470
      TabIndex        =   6
      Top             =   4410
      Width           =   4740
   End
   Begin VB.Label Label4 
      BackColor       =   &H0095C5E1&
      Caption         =   "Instrucciones:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   5
      Top             =   4410
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   7770
      TabIndex        =   4
      Top             =   3885
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6405
      TabIndex        =   3
      Top             =   3885
      Width           =   1275
   End
End
Attribute VB_Name = "RECUNOBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub TOTALI()
   SUDEBE# = 0: SUHABE# = 0
   For KU& = 1 To List1.ListCount
     SUDEBE# = SUDEBE# + Val(Mid$(List1.List(KU& - 1), 57, 14))
     SUHABE# = SUHABE# + Val(Mid$(List1.List(KU& - 1), 71, 14))
   Next KU&
   '
   Label1 = TRIM$(FORMATNUM(SUDEBE#, "F14.2"))
   Label2 = TRIM$(FORMATNUM(SUHABE#, "F14.2"))
End Sub

Private Sub Command1_Click()
   If Abs(Val(Label1) - Val(Label2)) >= 0.005 Then
      CUECOX$ = TRIM$(Label9)
      If CUECOX$ <> "" Then
         CuentaCo% = CUEINT%(CUECOX$)
         If CuentaCo% > 0 Then GoTo 25
      End If
      '
      Call MENSERR(24, "Imposible Registrar Comprobante no Balanceado.")
      Exit Sub
   End If
   '
25 If COMALTER%("Está Seguro de Registrar\el Comprobante Así Corregido ?\\No, Conservar Anterior\Si, Registrar") = 2 Then
       Label6 = "OK"
       Hide
     Else
       Unload Me
   End If
   '
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Command9_Click()
    Call SELECHO("ECUECOT")
    PRICUEX$ = VALACT1$("ECUECOT")
    If PRICUEX$ <> "" Then
        Label9 = PRICUEX$
    End If
End Sub

Private Sub Form_Load()
    If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
    End If

    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
End Sub

Private Sub Frame1_DblClick()
   Label9 = ""
   Label7 = ""
End Sub

Private Sub Label9_Change()
    Label7 = ECOARCH("ECUECOT", Label9)
End Sub

Private Sub Label9_DblClick()
   Call Command9_Click
End Sub

Private Sub List1_DblClick()
   KX& = List1.ListIndex
   List2.AddItem List1.TEXT
   List1.RemoveItem (KX&)
   Call TOTALI
End Sub
