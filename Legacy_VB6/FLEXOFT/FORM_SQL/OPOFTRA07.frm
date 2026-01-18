VERSION 5.00
Begin VB.Form OPOFTRA07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Trazabilidad - Configuración"
   ClientHeight    =   2580
   ClientLeft      =   -15
   ClientTop       =   240
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "IMPRESION DE ETIQUETA DE CIERRE DE O-FABRICACION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   4995
      Begin VB.CommandButton Command6 
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
         Height          =   690
         Left            =   4095
         Picture         =   "OPOFTRA07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   315
         Width           =   650
      End
      Begin VB.TextBox Text2 
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
         Left            =   240
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   " "
         Top             =   690
         Width           =   3750
      End
      Begin VB.CommandButton Command4 
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
         Left            =   3780
         TabIndex        =   11
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text1 
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
         Left            =   2415
         TabIndex        =   10
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario"
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
         Left            =   0
         TabIndex        =   14
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   5280
      ScaleHeight     =   4305
      ScaleWidth      =   1140
      TabIndex        =   6
      Top             =   0
      Width           =   1170
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   15
         Top             =   2160
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   16
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
         Height          =   690
         Left            =   105
         Picture         =   "OPOFTRA07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   260
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "OPOFTRA07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   950
         Width           =   650
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      Caption         =   "IMPRESION DE SOLICITUD DE MATERIALES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4995
      Begin VB.TextBox Text4 
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
         Left            =   2415
         TabIndex        =   4
         Text            =   " "
         Top             =   315
         Width           =   1380
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
         Height          =   300
         Left            =   3780
         TabIndex        =   3
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text5 
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
         Left            =   210
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   690
         Width           =   3750
      End
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
         Height          =   690
         Left            =   4095
         Picture         =   "OPOFTRA07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   315
         Width           =   650
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario"
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
         Left            =   0
         TabIndex        =   5
         Top             =   420
         Width           =   2325
      End
   End
End
Attribute VB_Name = "OPOFTRA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    Call GRACONTROL("", "FORCONSU", Text4.TEXT)
    '
    Call GRACONTROL("", "ETICIEOF", Text1.TEXT)
    '
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command3_Click()
    Call CommSel("INDIFRM", Command3, Text4)
End Sub

Private Sub Command4_Click()
Call CommSel("INDIFRM", Command4, Text1)
End Sub

Private Sub Command5_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text4.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command6_Click()
   Call CONECRUN("FLFORMS/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
End Sub

Private Sub Form_Load()
    Text4.TEXT = CONTROL$("", "FORCONSU")
    Text1.TEXT = CONTROL$("", "ETICIEOF")
End Sub

Private Sub Text1_Change()
    Text2 = ECOARCH$("INDIFRM", Text1)

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

