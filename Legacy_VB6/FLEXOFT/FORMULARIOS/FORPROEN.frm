VERSION 5.00
Begin VB.Form FORPROEN 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fechas Proyectadas de Entrega"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "O.K."
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
      Left            =   3360
      Picture         =   "FORPROEN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Genera y Emite Reporte"
      Top             =   2100
      Width           =   750
   End
   Begin VB.CommandButton Command1 
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
      Height          =   750
      Left            =   3360
      Picture         =   "FORPROEN.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Cierra y Abandona"
      Top             =   315
      Width           =   750
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
      Left            =   3360
      Picture         =   "FORPROEN.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1155
      Width           =   750
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   210
      TabIndex        =   5
      Top             =   1995
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
         Left            =   1680
         TabIndex        =   7
         Top             =   420
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
         Top             =   420
         Value           =   -1  'True
         Width           =   2500
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
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
         Top             =   1275
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
         Top             =   990
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
         Top             =   705
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
         Top             =   420
         Value           =   -1  'True
         Width           =   2500
      End
   End
End
Attribute VB_Name = "FORPROEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Hide
End Sub

Private Sub Command2_Click()
    '
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

