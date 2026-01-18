VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form SELITPED 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anulacion de Items de Pedidos"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11790
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2730
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   6
      Top             =   630
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   11000
      ScaleHeight     =   3105
      ScaleWidth      =   915
      TabIndex        =   3
      Top             =   0
      Width           =   950
      Begin VB.CommandButton Command19 
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
         Height          =   570
         Left            =   100
         Picture         =   "SELITPED.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Anula Items Seleccionados."
         Top             =   2280
         Width           =   570
      End
      Begin VB.CommandButton Command18 
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
         Height          =   570
         Left            =   100
         Picture         =   "SELITPED.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancelar"
         Top             =   400
         Width           =   570
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10905
      TabIndex        =   1
      Top             =   0
      Width           =   10965
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "  Código            Descripción                               Q.Ped      Q.Ent    Q.Saldo"
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   0
         TabIndex        =   2
         ToolTipText     =   "Haga Doble-Click para Blanquear Pre-Carga y Habilitar Carga Manual"
         Top             =   105
         Width           =   10755
      End
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00E0E0E0&
      Height          =   2610
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   440
      Width           =   10965
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "SELITPED.frx":0614
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "SELITPED.frx":0848
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Line Line5 
      X1              =   10970
      X2              =   10970
      Y1              =   0
      Y2              =   3200
   End
End
Attribute VB_Name = "SELITPED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command18_Click()
    List2.Clear
    Unload Me
End Sub

Private Sub Command19_Click()
    Hide
End Sub

Private Sub Form_Load()
    
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture1)

End Sub

Private Sub List2_DblClick()
    ILISTA& = List2.ListIndex
    TTXX$ = List2.List(ILISTA&)
    List2.RemoveItem (ILISTA&)
    STATX$ = TRIM$(Mid$(TTXX$, 71, 6))
    If STATX$ = "" Then
         STATX$ = "anul."
       ElseIf STATX$ = "anul." Then
         STATX$ = "      "
    End If
    Call REPLA(TTXX$, STATX$, 71, 6)
    List2.AddItem TTXX$, ILISTA&
    List2.ListIndex = ILISTA&
End Sub
