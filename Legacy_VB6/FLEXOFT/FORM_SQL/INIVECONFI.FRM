VERSION 5.00
Begin VB.Form INIVECONFI 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0F0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verifica Controlador Fiscal"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5265
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Verifica Numeración"
      Height          =   225
      Left            =   210
      TabIndex        =   5
      Top             =   2730
      Width           =   3900
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   4575
      Left            =   4305
      ScaleHeight     =   4515
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Picture         =   "INIVECONFI.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   810
         Width           =   600
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
         Height          =   600
         Left            =   120
         Picture         =   "INIVECONFI.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   210
         Width           =   600
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   2175
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "INIVECONFI.frx":0454
         Top             =   2520
         Width           =   1515
      End
   End
   Begin VB.Line Line3 
      X1              =   -105
      X2              =   5040
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Proxima N.Credito 'B':"
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
      Left            =   -840
      TabIndex        =   15
      Top             =   2205
      Width           =   3480
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2730
      TabIndex        =   14
      Top             =   2100
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Proxima N.Credito 'A':"
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
      Left            =   -840
      TabIndex        =   13
      Top             =   1785
      Width           =   3480
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2730
      TabIndex        =   12
      Top             =   1680
      Width           =   1380
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Proxima Factura / ND 'B':"
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
      Left            =   -840
      TabIndex        =   11
      Top             =   1155
      Width           =   3480
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2730
      TabIndex        =   10
      Top             =   1050
      Width           =   1380
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2730
      TabIndex        =   9
      Top             =   630
      Width           =   1380
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2730
      TabIndex        =   8
      Top             =   210
      Width           =   750
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Proxima Factura / ND 'A':"
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
      Left            =   -840
      TabIndex        =   7
      Top             =   735
      Width           =   3480
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Punto de Venta:"
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
      Left            =   -840
      TabIndex        =   6
      Top             =   315
      Width           =   3480
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   5145
      Y1              =   1575
      Y2              =   1575
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5880
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "INIVECONFI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
   Call FORHAS320.INICIA
   '
   Call FORHAS320.TOMANUMDELCONFIS("FC", "A", PUNTODEVENTA$, NUMECOMP$)
   Label5 = PUNTODEVENTA$
   Label6 = NUMECOMP$
   '
   Call FORHAS320.TOMANUMDELCONFIS("FC", "B", PUNTODEVENTA$, NUMECOMP$)
   Label7 = NUMECOMP$
   '
   Call FORHAS320.TOMANUMDELCONFIS("NC", "A", PUNTODEVENTA$, NUMECOMP$)
   Label1 = NUMECOMP$
   '
   Call FORHAS320.TOMANUMDELCONFIS("NC", "B", PUNTODEVENTA$, NUMECOMP$)
   Label9 = NUMECOMP$

End Sub

Private Sub Form_Load()
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   Set FORFACPRINT = FORFASQL
   '
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call FINAL("")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call FINAL("")
End Sub
