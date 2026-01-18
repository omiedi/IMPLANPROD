VERSION 5.00
Begin VB.Form OPTRAZ04 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Registro de Trazabilidad- Configuración"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6450
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
   ScaleHeight     =   2670
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   2850
      Left            =   5565
      ScaleHeight     =   2790
      ScaleWidth      =   1110
      TabIndex        =   25
      Top             =   0
      Width           =   1170
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   90
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   28
         Top             =   1750
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   610
         Left            =   105
         Picture         =   "OPTRAZ04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   260
         Width           =   645
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
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
         Left            =   105
         Picture         =   "OPTRAZ04.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   870
         Width           =   645
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "OPTRAZ04.frx":0454
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   2080
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Selección de Partidas a O/F's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   22
      Top             =   1575
      Width           =   5160
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Automática por Criterio FIFO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2835
         TabIndex        =   24
         Top             =   360
         Width           =   2040
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Asignación Manual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   840
         TabIndex        =   23
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Impresión de Registro de Trazabilidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   210
      TabIndex        =   16
      Top             =   210
      Width           =   5160
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   4095
         Picture         =   "OPTRAZ04.frx":2376
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   315
         Width           =   750
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   3750
      End
      Begin VB.CommandButton Command3 
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
         Height          =   360
         Left            =   3780
         TabIndex        =   18
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2415
         TabIndex        =   17
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   0
         TabIndex        =   21
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Origen de Stocks Remotos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   210
      TabIndex        =   7
      Top             =   2940
      Width           =   5160
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   735
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1995
         Width           =   4230
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   735
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1470
         Width           =   4230
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   735
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   945
         Width           =   4230
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   735
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   420
         Width           =   4230
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-4:"
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
         Index           =   3
         Left            =   0
         TabIndex        =   15
         Top             =   2100
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-3:"
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
         Index           =   2
         Left            =   0
         TabIndex        =   13
         Top             =   1575
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-2:"
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
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   1050
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-1:"
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
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   525
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Destino Archivo de Transferencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   4
      Top             =   6720
      Width           =   5160
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   735
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   420
         Width           =   4230
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta:"
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
         Left            =   0
         TabIndex        =   6
         Top             =   525
         Width           =   645
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Identificador de Sede"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   1
      Top             =   5565
      Width           =   5160
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3780
         MaxLength       =   8
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   420
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código Identificador de Sede:"
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
         Left            =   105
         TabIndex        =   3
         Top             =   525
         Width           =   3480
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
      Picture         =   "OPTRAZ04.frx":2680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3990
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   5040
      Picture         =   "OPTRAZ04.frx":298A
      Top             =   6950
      Width           =   2010
   End
End
Attribute VB_Name = "OPTRAZ04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '\\\OT-06-0476-RG-07/12/06///    '
    'Call GRACONTROL("INDIFRM", "FORTRAZA", Text4.TEXT)
    Call GRACONTROL("TRAZABI", "FORTRAZA", Text4.TEXT)
    '
    MODOTRAZ$ = "MANUAL"
    If Option2.Value = True Then MODOTRAZ$ = "AUTOMA"
    Call GRACONTROL("TRAZABI", "MODOTRAZ", MODOTRAZ$)
    '
    Call CIERRARCH("*.*")
    Hide
    '\\\OT-06-0476-RG-FIN///    '
End Sub

Private Sub Command3_Click()
    Call CommSel("INDIFRM", Command3, Text4)
End Sub

Private Sub Command5_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text4.TEXT), "Formularios de Impresion")
End Sub

Private Sub Form_Load()
    Text4.TEXT = CONTROL$("TRAZABI", "FORTRAZA")
    If CONTROL$("TRAZABI", "MODOTRAZ") = "AUTOMA" Then
       Option2.Value = True
    End If
End Sub

Private Sub Text4_Change()
    Text5 = ECOARCH$("INDIFRM", Text4)
End Sub

Private Sub Text4_DblClick()
   Call Command3_Click
End Sub

Private Sub Text5_Click()
   Text4.SetFocus
End Sub
