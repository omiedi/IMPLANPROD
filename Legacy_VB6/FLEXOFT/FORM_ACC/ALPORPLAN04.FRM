VERSION 5.00
Begin VB.Form ALPORPLA04 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alta de Código de Artículo por Plano - Posición"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   8925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
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
      Height          =   330
      Left            =   4620
      Picture         =   "ALPORPLAN04.frx":0000
      TabIndex        =   15
      Top             =   720
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
      Height          =   330
      Left            =   4620
      Picture         =   "ALPORPLAN04.frx":030A
      TabIndex        =   13
      Top             =   240
      Width           =   175
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   4530
      Left            =   7920
      ScaleHeight     =   4470
      ScaleWidth      =   1020
      TabIndex        =   7
      Top             =   0
      Width           =   1080
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   555
         TabIndex        =   11
         Top             =   3600
         Width           =   615
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command8 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "ALPORPLAN04.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Acceso a Maestro de Artículos"
         Top             =   1920
         Width           =   640
      End
      Begin VB.CommandButton Command3 
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
         Left            =   120
         Picture         =   "ALPORPLAN04.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   100
         Width           =   640
      End
      Begin VB.CommandButton BOTOK 
         Caption         =   "Ok"
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
         Left            =   120
         Picture         =   "ALPORPLAN04.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Alta del Nuevo Código"
         Top             =   790
         Width           =   640
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -240
         Picture         =   "ALPORPLAN04.frx":0D72
         Top             =   3900
         Width           =   1515
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      TabIndex        =   5
      Top             =   1155
      Width           =   6105
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      MaxLength       =   15
      TabIndex        =   0
      ToolTipText     =   "Introduzca el Nuevo Código."
      Top             =   600
      Width           =   2415
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
      Height          =   330
      Left            =   4200
      MaxLength       =   15
      TabIndex        =   14
      Top             =   240
      Width           =   495
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
      Height          =   330
      Left            =   4200
      MaxLength       =   15
      TabIndex        =   16
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Codigo Plano/Posición"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Denominación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   -840
      TabIndex        =   6
      Top             =   1250
      Width           =   2430
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3480
      TabIndex        =   4
      Top             =   735
      Width           =   540
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
      Height          =   330
      Left            =   4935
      TabIndex        =   3
      Top             =   750
      Width           =   2850
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "U.Med."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3120
      TabIndex        =   2
      Top             =   315
      Width           =   1065
   End
   Begin VB.Label Label3 
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
      Height          =   330
      Left            =   4935
      TabIndex        =   1
      Top             =   210
      Width           =   2850
   End
End
Attribute VB_Name = "ALPORPLA04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BOTOK_Click()
        '
    BOTOK.Enabled = False
    '
'    If Len(Text1) < 3 Or Len(Text1) > 15 Then
'      Call COMUNI("Código de Articulo No Válido")
'      GoTo 99
'    End If
'    '
'    If TRIM$(Label3) = "" Then
'        Call COMUNI("Falta Unidad de Medida del Articulo")
'        GoTo 99
'    End If
'    '
'    If TRIM$(Label5) = "" Then
'        Call COMUNI("Falta Tipo de Material del Articulo")
'        GoTo 99
'    End If
'    '
'    If TRIM$(Text2) = "" Then
'      Call COMUNI("Falta Descripción del Artículo")
'      GoTo 99
'    End If
    '
    Screen.MousePointer = 11
    CODINUE$ = TRIM$(Text1)
    DESCRINUE$ = Text2
    TIMA% = Val(Text4)
    UNID$ = TRIM$(Text3)
    '
    Call AGRECOD(CODINUE$, DESCRINUE$, TIMA%, UNID$, OKX%)
    '/////////***********/////////////////
    If OKX% < 1 Then
      Call MENSERR(24, "El Alta de Codigo\no pudo Completarse.\  \Ejecute Control de Maestro\y Consulte a su Soporte Técnico.")
    Else
      Text3 = ""
      Text1 = ""
      Text2 = ""
      Text4 = ""
      Label3 = ""
      Label5 = ""
      Hide
    End If
    '
99  Screen.MousePointer = 1
    BOTOK.Enabled = True
    '
End Sub

Private Sub Command1_Click()
   Call SELECHO("UMEDIDA")
   VALONUE$ = VALACT1$("UMEDIDA")
   If VALONUE$ <> "" Then
      Text3.TEXT = VALONUE$
   End If
End Sub

Private Sub Command2_Click()
     Call SELECHO("TATIMAT")
   VALONUE$ = VALACT1$("TATIMAT")
   If VALONUE$ <> "" Then
      Text4.TEXT = VALONUE$
   End If
End Sub

Private Sub Command3_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    AMASTO00.Show 1
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    '
    X$ = REGSEL$("TAFAMIL")
    '
    If Len(X$) > 4 Then
      Text1 = Mid$(X$, 1, 2)
    End If
    '
End Sub



Private Sub Label7_Change()
    '
    X$ = ECOARCH("TATIMAT", Chr(Val(Label7)))
    Label5 = TRIM$(X$)
    '
End Sub

Private Sub Text1_GotFocus()
    '
    Command1.SetFocus
    '
End Sub
Private Sub Text2_GotFocus()
    '
    Command1.SetFocus
    '
End Sub
Private Sub Text1_DblClick()
    '
    'Command9_Click
    '
End Sub

Private Sub Text3_Change()
   '
  
   Label3.Caption = ECOARCH$("UMEDIDA", Text3.TEXT)
   '
End Sub

Private Sub Text4_Change()
   VALONU = Val(Text4.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text6.TEXT = "": Exit Sub
   End If
   Label5.Caption = ECOARCH$("TATIMAT", Chr$(VALONU))
End Sub
