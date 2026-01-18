VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form DESHASMES 
   BackColor       =   &H00E0E0D0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección Mes Desde - Hasta"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   300
      Index           =   0
      Left            =   3360
      ScaleHeight     =   240
      ScaleWidth      =   1110
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   2850
      Left            =   4680
      ScaleHeight     =   2790
      ScaleWidth      =   1110
      TabIndex        =   5
      Top             =   0
      Width           =   1170
      Begin VB.CommandButton Command15 
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
         Height          =   555
         Left            =   105
         Picture         =   "DESHASMES.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Cancelar Cambio de Cotización"
         Top             =   90
         Width           =   600
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
         Height          =   555
         Left            =   105
         Picture         =   "DESHASMES.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Aprobar Cambio de Cotización"
         Top             =   645
         Width           =   600
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0D0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4650
      Begin VB.CommandButton Command2 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   17
         Top             =   900
         Width           =   175
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
         Height          =   285
         Left            =   4305
         TabIndex        =   11
         Top             =   465
         Width           =   175
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
         Height          =   285
         Left            =   2025
         TabIndex        =   9
         Top             =   465
         Width           =   175
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   18
         Top             =   900
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   690
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2610
         TabIndex        =   14
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   340
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2475
         TabIndex        =   4
         Top             =   240
         Width           =   570
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5600
      Picture         =   "DESHASMES.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3990
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "DESHASMES.frx":091E
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "DESHASMES.frx":0B52
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   5040
      Picture         =   "DESHASMES.frx":0BCC
      Top             =   6950
      Width           =   2010
   End
End
Attribute VB_Name = "DESHASMES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DIADESDE%
Public DIAHASTA%
Private Sub Command1_Click()
   Command1.Enabled = False
   Call SELMESA(DES%, HAS%, PERIO$)
   DIAHASTA% = HAS%
   Label5 = PERIO$
   Command1.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call SELMESA(DES%, HAS%, PERIO$)
   DIADESDE% = DES%
   Label3 = PERIO$
   Command12.Enabled = True
End Sub

Private Sub Command14_Click()
   Hide
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   Label3.Caption = "": Label5.Caption = "": Label7.Caption = "": Label9.Caption = ""
   DIADESDE% = 0
   DIAHASTA% = 0
   Command15.Enabled = True
End Sub

Private Sub Command2_Click()

    Call SELECHO("DEUDOR12")
    NCLI = XVALO(TRIM$(VALACT1$("DEUDOR12")))
    If NCLI > 0 Then
        Label7.Caption = TRIM$(VALACT1$("DEUDOR12"))
        Label9.Caption = TRIM$(VALACT2$("DEUDOR12"))
    Else
        Label7.Caption = ""
        Label9.Caption = ""
    End If
    
End Sub

Private Sub Form_Load()
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me)
End Sub
