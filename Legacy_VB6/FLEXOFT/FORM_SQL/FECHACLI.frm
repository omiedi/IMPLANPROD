VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form FECHACLI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pendientes de Cobranza"
   ClientHeight    =   2925
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1680
      OleObjectBlob   =   "FECHACLI.frx":0000
      Top             =   1080
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rango de Fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1680
         Width           =   765
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "Legajo:"
         Top             =   1320
         Width           =   765
      End
      Begin VB.CommandButton Command2 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1005
         TabIndex        =   8
         Top             =   1680
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1680
         Width           =   3915
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3900
         TabIndex        =   3
         Top             =   2160
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   495
         Index           =   0
         Left            =   45
         OleObjectBlob   =   "FECHACLI.frx":0234
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "FECHACLI.frx":029E
         TabIndex        =   2
         Top             =   2160
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker Desde 
         Height          =   330
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   128778241
         CurrentDate     =   41123
      End
      Begin MSComCtl2.DTPicker Hasta 
         Height          =   330
         Left            =   3840
         TabIndex        =   5
         Top             =   360
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   128778241
         CurrentDate     =   41123
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   495
         Index           =   1
         Left            =   2880
         OleObjectBlob   =   "FECHACLI.frx":02FC
         TabIndex        =   6
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "FECHACLI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Hide
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call ABRECONEXION
    '
    If ALUMONOSEL > 0 Then
      Text4 = XVALO(RESP_ZELE_VECT(1))
    End If
    '
99  Command2.Enabled = True
End Sub
Private Sub Form_Load()

    Call APLICAR_SKIN1(Me, App.Path & "\SKINS\dogmax.skn")


End Sub

Private Sub Text4_Change()
   LEGAJO& = XVALO(Text4)
   CONDI$ = "NUME_CLI = " & LEGAJO&
   '
   SQLX$ = "SELECT * FROM DEUDOR12 WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " + CONDI$
   Set ALUMTMP = FLEXCONN.Execute(SQLX$)
   With ALUMTMP
        Text1 = SAFETEXT$(!RASO_CLI)
   End With
   ALUMTMP.Close
   Set ALUMTMP = Nothing

End Sub
