VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CONSULTAS1 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consultas"
   ClientHeight    =   7695
   ClientLeft      =   525
   ClientTop       =   2505
   ClientWidth     =   12630
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   12630
   ShowInTaskbar   =   0   'False
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "CONSULTAS1.frx":0000
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Imprimir"
      Height          =   690
      Left            =   5400
      Picture         =   "CONSULTAS1.frx":006A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   7095
   End
   Begin VB.CommandButton Command98 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   120
      Picture         =   "CONSULTAS1.frx":5C7C
      TabIndex        =   1
      Top             =   6840
      Width           =   3075
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5640
      OleObjectBlob   =   "CONSULTAS1.frx":5DC6
      Top             =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   10821
      _Version        =   393216
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "CONSULTAS1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command25_Click()
   Command25.Enabled = False
   'valido
   If MSF1.Rows <= 1 Then
      MsgBox "GRILLA DE IMPRESIÓN VACIA"
      GoTo 99
   End If
   '
   Set GRILLA_IMPRESION = MSF1
   PREIMPRE.txtEncabezado1.Text = SkinLabel8.Caption

   PREIMPRE.Show 1
   '
99 Command25.Enabled = True
End Sub

Private Sub Command98_Click()
   Command98.Enabled = False
   Unload Me
   Command98.Enabled = True
End Sub

Private Sub Form_Load()
'
    Call APLICAR_SKIN(Me, App.Path & "\SKINS\dogmax.skn")
    Me.Top = mdiprincipal.Top + mdiprincipal.Height / 2 - Me.Height / 2
    Me.Left = mdiprincipal.Left + mdiprincipal.Width / 2 - Me.Width / 2

End Sub
