VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPIMPUTA 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Re-Imputación Débitos / Créditos"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E2D3C0&
      Caption         =   "Gastos Bancarios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   420
      TabIndex        =   15
      Top             =   1560
      Width           =   3480
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   13
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.OptionButton Option4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E2D3C0&
      Caption         =   "Intereses por Mora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   420
      TabIndex        =   12
      Top             =   1200
      Width           =   3480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   3405
      Left            =   4200
      ScaleHeight     =   3375
      ScaleWidth      =   1560
      TabIndex        =   9
      Top             =   0
      Width           =   1590
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
         Height          =   690
         Left            =   105
         Picture         =   "OPIMPUTA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Aplica Imputación Contable Según Selección"
         Top             =   850
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Height          =   690
         Left            =   105
         Picture         =   "OPIMPUTA.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Mantiene Imputación Contable General"
         Top             =   150
         Width           =   650
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E2D3C0&
      Caption         =   "IMPUTACIÓN CONFIGURADA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   1155
      Left            =   100
      TabIndex        =   3
      Top             =   1920
      Width           =   3930
      Begin VB.TextBox Text16 
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
         Left            =   840
         TabIndex        =   5
         Top             =   345
         Width           =   1170
      End
      Begin VB.CommandButton Command12 
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
         Left            =   1995
         TabIndex        =   4
         Top             =   345
         Width           =   180
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   -3045
         TabIndex        =   8
         Top             =   1205
         Width           =   4110
      End
      Begin VB.Label Label21 
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
         Height          =   330
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   420
         Width           =   615
      End
      Begin VB.Label Label16 
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
         Height          =   300
         Left            =   840
         TabIndex        =   6
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.OptionButton Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E2D3C0&
      Caption         =   "Débitos / Créditos Varios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   420
      TabIndex        =   2
      Top             =   840
      Width           =   3480
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E2D3C0&
      Caption         =   "Diferencia de Cambio o Cotización"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   420
      TabIndex        =   1
      Top             =   480
      Width           =   3480
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E2D3C0&
      Caption         =   "Cheques Rechazados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   420
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   2640
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "OPIMPUTA.frx":0614
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "OPIMPUTA.frx":0848
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "OPIMPUTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'\\\OT-06-0403-MH-19/09/06///
   On Error Resume Next
   CUEVE$ = FORFACCS.Text16.TEXT
   If Err Then
     CUEVE$ = FORFASQL.Text16.TEXT
   End If
   On Error GoTo 0
   QQQ% = CUECOINT%(CUEVE$)
   Call COMUNI("El Documento se Contabiliza\en la Cuenta Contable\'" + TRIM$(CUEVE$) + " - " + TRIM$(DENOCUE$(QQQ%)) + "'")
   Hide
End Sub
'\\\OT-06-0403-MH-FIN///

'\\\OT-06-0403-MH-18/09/06///
Private Sub Command12_Click()

   If ULTREG("CTASFACT") > 0 Then
      Call SELECHO("CTASFACT")
      VDEVU$ = TRIM$(VALACT1$("CTASFACT"))
      If VDEVU$ <> "" Then
        Text16 = CODCUE$(XVALO(VDEVU$))
      End If
   Else
      Call SELECHO("ECUECON")
      VDEVU$ = TRIM$(VALACT1$("ECUECON"))
      If VDEVU$ <> "" Then
        Text16 = VDEVU$
      End If
   End If
    
End Sub

Private Sub Command14_Click()
   CUEVE$ = Text16.TEXT
   QQQ% = CUECOINT%(CUEVE$)
   If TRIM$(CUEVE$) <> "" Then
      If QQQ% > 0 Then
         Call COMUNI("El Documento se Contabiliza\en la Cuenta Contable\'" + TRIM$(CUEVE$) + " - " + TRIM$(DENOCUE$(QQQ%)) + "'")
         On Error Resume Next
         FORFACCS.Text16 = CUEVE$
         FORFASQL.Text16 = CUEVE$
         On Error GoTo 0
         Hide
         Exit Sub
      End If
   End If
   Call MENSERR(24, "Cuenta Invalida")
99 End Sub
'
Private Sub Form_Load()
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture1)
    CUEVE$ = Control$("SVD451S", "CUEVE-8")
    Text16.TEXT = CUEVE$
End Sub
'
Sub Option1_Click()
    CUEVE$ = Control$("SVD451S", "CUEVE-8")
    Text16.TEXT = CUEVE$
End Sub

 Sub Option2_Click()
    CUEVE$ = Control$("SVD451S", "CUEVE-9")
    Text16.TEXT = CUEVE$
End Sub

 Sub Option3_Click()
    CUEVE$ = Control$("SVD451S", "CUEVE-11")
    Text16.TEXT = CUEVE$
End Sub

Sub Option4_Click()
    CUEVE$ = Control$("SVD451S", "CUEVE-14")
    Text16.TEXT = CUEVE$
End Sub

Private Sub Option5_Click()
    CUEVE$ = Control$("SVD451S", "CUEVE-16")
    Text16.TEXT = CUEVE$
End Sub

Private Sub Text16_Change()
   Label16 = Left$(ECOARCH("ECUECON", Text16), 20)
End Sub
