VERSION 5.00
Begin VB.Form STPCALPR09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema de Calificación de Proveedores"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5610
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
   ScaleHeight     =   2055
   ScaleWidth      =   5610
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Categoria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   180
      TabIndex        =   6
      Top             =   1155
      Width           =   4290
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Left            =   3300
         TabIndex        =   12
         Top             =   345
         Width           =   690
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
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
         Left            =   1935
         TabIndex        =   10
         Top             =   345
         Width           =   690
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
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
         Left            =   630
         TabIndex        =   8
         Top             =   345
         Width           =   690
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2955
         TabIndex        =   11
         Top             =   405
         Width           =   285
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1590
         TabIndex        =   9
         Top             =   405
         Width           =   285
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   270
         TabIndex        =   7
         Top             =   405
         Width           =   285
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Ponderacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   180
      TabIndex        =   1
      Top             =   240
      Width           =   4305
      Begin VB.TextBox Text2 
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
         Left            =   3315
         TabIndex        =   5
         Top             =   330
         Width           =   690
      End
      Begin VB.TextBox Text1 
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
         Left            =   1080
         TabIndex        =   4
         Top             =   330
         Width           =   690
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cumplimiento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2115
         TabIndex        =   3
         Top             =   375
         Width           =   1185
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Calidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   60
         TabIndex        =   2
         Top             =   405
         Width           =   960
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   2940
      Left            =   4680
      ScaleHeight     =   2880
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   0
      Width           =   1035
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
         Height          =   750
         Left            =   120
         Picture         =   "STPCALPR09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Registra y Guarda la Configuración"
         Top             =   240
         Width           =   680
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -285
         Picture         =   "STPCALPR09.frx":030A
         Top             =   1480
         Width           =   1515
      End
   End
End
Attribute VB_Name = "STPCALPR09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command2_Click()
   Command2.Enabled = False
   Call GRACONTROL("CALIPROV", "ESCACALI", TRIM$(Str$(ESCACALI)))
   'Call GRACONTROL("CALIPROV", "PUMAXCAL", TRIM$(Label2))
   '
   SUMA# = XVALO(Text1) + XVALO(Text2)
   If SUMA# <> 100 Then
      COMUNI ("La Suma de Ponderacion Debe ser Igual a 100")
      GoTo 99
   End If
   '
   Call GRACONTROL("CALIPROV", "PONDECAL", TRIM$(Text1))
   Call GRACONTROL("CALIPROV", "PONDECUM", TRIM$(Text2))
   '
   Call GRACONTROL("CALIPROV", "LIMCAT-A", TRIM$(Text3))
   Call GRACONTROL("CALIPROV", "LIMCAT-B", TRIM$(Text4))
   Call GRACONTROL("CALIPROV", "LIMCAT-C", TRIM$(Text5))
99
   '
   Command2.Enabled = True
End Sub
'

Private Sub Form_Load()
   '
   Text1 = CONTROL("CALIPROV", "PONDECAL")
   Text2 = CONTROL("CALIPROV", "PONDECUM")
   Text3 = CONTROL("CALIPROV", "LIMCAT-A")
   Text4 = CONTROL("CALIPROV", "LIMCAT-B")
   Text5 = CONTROL("CALIPROV", "LIMCAT-C")
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub

Private Sub Text1_LostFocus()
Text1 = FORMATNUM(Text1, "F5." + CNTDC$)
End Sub

Private Sub Text2_LostFocus()
Text2 = FORMATNUM(Text2, "F5." + CNTDC$)
End Sub

Private Sub Text3_LostFocus()
Text3 = FORMATNUM(Text3, "F5.2")
End Sub

Private Sub Text4_LostFocus()
   Text4 = FORMATNUM(Text4, "F5.2")
End Sub

Private Sub Text5_LostFocus()
   Text5 = FORMATNUM(Text5, "F5.2")
End Sub
