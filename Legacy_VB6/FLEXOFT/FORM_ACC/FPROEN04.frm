VERSION 5.00
Begin VB.Form FPROEN04 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fechas Proyectadas de Entrega"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   3585
      Left            =   3360
      ScaleHeight     =   3525
      ScaleWidth      =   1005
      TabIndex        =   9
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   635
         Left            =   105
         Picture         =   "FPROEN04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Genera y Emite Reporte"
         Top             =   1155
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   11
         Top             =   1950
         Width           =   650
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
         Height          =   635
         Left            =   105
         Picture         =   "FPROEN04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   270
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "FPROEN04.frx":0454
         Top             =   2205
         Width           =   1515
      End
   End
   Begin VB.CommandButton Command29 
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
      Height          =   750
      Left            =   630
      Picture         =   "FPROEN04.frx":2376
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   3990
      Width           =   750
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   210
      TabIndex        =   5
      Top             =   1890
      Width           =   2955
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1575
         TabIndex        =   7
         Top             =   315
         Width           =   1245
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Pantalla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   420
         TabIndex        =   6
         Top             =   315
         Value           =   -1  'True
         Width           =   2500
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   2955
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Por Número de O/F"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   420
         TabIndex        =   4
         Top             =   1170
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Por Producto / Artículo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   420
         TabIndex        =   3
         Top             =   885
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Por Destino / Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   420
         TabIndex        =   2
         Top             =   600
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   420
         TabIndex        =   1
         Top             =   315
         Value           =   -1  'True
         Width           =   2500
      End
   End
End
Attribute VB_Name = "FPROEN04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Hide
End Sub

Private Sub Command2_Click()
    '
    Call GENORDEFDAT(1)
    If Option2(1).Value = True Then
        Screen.MousePointer = 11
        Call COPYSTRU("ORDEFABR", "ORDESOR")
        Call BLANARCH("ORDESOR")
        For U& = 1 To ULTREG&("ORDEFABR")
           XX$ = REGLEIDO$("ORDEFABR", U&)
           ST% = Asc(Mid$(XX$, 108, 1))
           If ST% > 0 Then
              If ST% < 3 Then
                 Call REGAPP("ORDESOR", XX$)
              End If
           End If
        Next U&
        Screen.MousePointer = 1
        '
        If Option1(0).Value = True Then Call FINAL("*SFEPRO0")
        If Option1(1).Value = True Then Call FINAL("*SFEPRO1")
        If Option1(2).Value = True Then Call FINAL("*SFEPRO2")
        If Option1(3).Value = True Then Call FINAL("*SFEPRO3")
        Exit Sub
    End If
    '
    If Option1(0).Value = True Then Call FINAL("*FEPROE0")
    If Option1(1).Value = True Then Call FINAL("*FEPROE1")
    If Option1(2).Value = True Then Call FINAL("*FEPROE2")
    If Option1(3).Value = True Then Call FINAL("*FEPROE3")
    '
End Sub

Private Sub Command29_Click()
    '
    Call HELPONLINE("FEPROEN")
    '
End Sub

