VERSION 5.00
Begin VB.Form CARFAPRO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga Datos Factura Proforma"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   4860
      Left            =   7170
      ScaleHeight     =   4800
      ScaleWidth      =   1020
      TabIndex        =   10
      Top             =   -15
      Width           =   1080
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
         Height          =   720
         Left            =   180
         Picture         =   "CARFAPRO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Aprueba, Graba e Imprime Presupuesto"
         Top             =   3795
         Width           =   660
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   11
         Top             =   315
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -315
            TabIndex        =   12
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -255
         Picture         =   "CARFAPRO.frx":030A
         Top             =   600
         Width           =   1515
      End
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   1815
      TabIndex        =   9
      Top             =   4050
      Width           =   3405
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   1815
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   3135
      Width           =   5235
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1815
      TabIndex        =   7
      Top             =   2460
      Width           =   3405
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Top             =   1800
      Width           =   3390
   End
   Begin VB.TextBox Text1 
      Height          =   1065
      Left            =   1815
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   345
      Width           =   5160
   End
   Begin VB.Label Label5 
      Caption         =   "Cond. Venta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   510
      TabIndex        =   5
      Top             =   4095
      Width           =   1155
   End
   Begin VB.Label Label4 
      Caption         =   "NCM:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1155
      TabIndex        =   4
      Top             =   3195
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Peso Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   630
      TabIndex        =   3
      Top             =   2520
      Width           =   1050
   End
   Begin VB.Label Label2 
      Caption         =   "Peso Neto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   660
      TabIndex        =   2
      Top             =   1890
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "Bultos / Embalaje:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   90
      TabIndex        =   0
      Top             =   405
      Width           =   1620
   End
End
Attribute VB_Name = "CARFAPRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OK As Boolean
Private Sub Command14_Click()
  '
  If XVALO(Text2) > 999999999 Then
     Call COMUNI("Peso Neto Supera Maximo Permitido")
     Exit Sub
  End If
  '
  If XVALO(Text3) > 999999999 Then
     Call COMUNI("Peso Total Supera Maximo Permitido")
     Exit Sub
  End If
  '
  OK = True
  Hide
End Sub

Private Sub Form_Load()
  OK = False
End Sub


