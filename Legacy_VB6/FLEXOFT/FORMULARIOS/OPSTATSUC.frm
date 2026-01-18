VERSION 5.00
Begin VB.Form OPSTATSUC 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones de Cálculo"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sucursales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   3885
      TabIndex        =   19
      Top             =   210
      Width           =   2625
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   105
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   20
         Top             =   315
         Width           =   2385
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   210
      TabIndex        =   12
      Top             =   2205
      Width           =   7260
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Seleccionados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   210
         TabIndex        =   16
         Top             =   315
         Width           =   3690
         Begin VB.CommandButton Command4 
            Caption         =   "Quitar Todos"
            Height          =   225
            Left            =   210
            TabIndex        =   18
            Top             =   1575
            Width           =   3240
         End
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1320
            Left            =   210
            Sorted          =   -1  'True
            TabIndex        =   17
            Top             =   315
            Width           =   3270
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Excluidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   3990
         TabIndex        =   13
         Top             =   315
         Width           =   3060
         Begin VB.CommandButton Command6 
            Caption         =   "Incluir Todos"
            Height          =   225
            Left            =   210
            TabIndex        =   15
            Top             =   1575
            Width           =   2610
         End
         Begin VB.ListBox List2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1320
            Left            =   210
            Sorted          =   -1  'True
            TabIndex        =   14
            Top             =   315
            Width           =   2640
         End
      End
   End
   Begin VB.CommandButton Command2 
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
      Height          =   850
      Left            =   6615
      Picture         =   "OPSTATSUC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   290
      Width           =   855
   End
   Begin VB.CommandButton Command3 
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
      Height          =   850
      Left            =   6615
      Picture         =   "OPSTATSUC.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Inicia Proceso de Cálculo Estadístico"
      Top             =   1290
      Width           =   855
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Meses"
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
      Left            =   210
      TabIndex        =   4
      Top             =   1290
      Width           =   3585
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "24"
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
         Left            =   2205
         TabIndex        =   8
         Top             =   400
         Width           =   600
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "12"
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
         Left            =   1470
         TabIndex        =   7
         Top             =   400
         Width           =   600
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "8"
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
         Left            =   840
         TabIndex        =   6
         Top             =   400
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "6"
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
         Left            =   210
         TabIndex        =   5
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2835
         TabIndex        =   9
         Top             =   255
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Considera Hasta"
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
      TabIndex        =   0
      Top             =   210
      Width           =   3585
      Begin VB.CommandButton Command5 
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
         Height          =   280
         Left            =   3150
         TabIndex        =   1
         Top             =   420
         Width           =   175
      End
      Begin VB.Label MESAN 
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   1260
         TabIndex        =   3
         Top             =   420
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Mes y Año:"
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
         TabIndex        =   2
         Top             =   480
         Width           =   1065
      End
   End
End
Attribute VB_Name = "OPSTATSUC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
   Label1 = ""
   Unload Me
End Sub

Private Sub Command3_Click()
   Hide
End Sub

Private Sub Command4_Click()
   For U& = 1 To List1.ListCount
     List2.AddItem List1.List(U& - 1)
   Next U&
   List1.Clear
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
      MESAN = MESASEL$
      MESAN.Refresh
   End If
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   For U& = 1 To List2.ListCount
     List1.AddItem List2.List(U& - 1)
   Next U&
   List2.Clear
End Sub

Private Sub Form_Load()
   Call CARSELMESA
   MESAN = Left$(REGLEIDO$("SELMESA", 4), 14)
   '
   List1.Clear: List2.Clear: List3.Clear
   For U& = 1 To ULTREG&("DEUDOR1")
     X$ = REGLEIDO$("DEUDOR1", U&)
     List2.AddItem Mid$(X$, 3, 30) + Space$(34) + TEXTBIN$(Left$(X$, 2), "I8")
   Next U&
   List2.ListIndex = (-1)
   '
   For U& = 1 To ULTREG&("TADEPOSI")
     X$ = REGLEIDO$("TADEPOSI", U&)
     ABREVI$ = TRIM$(Mid$(X$, 22, 8))
     If ABREVI$ <> "" Then
       IDEPO% = Asc(Left$(X$, 1))
       TTXX$ = AJUSTD$(Str$(IDEPO%), 3) + "  " + Mid$(X$, 2, 20) + Space$(39) + ABREVI$
       List3.AddItem TTXX$
     End If
   Next U&
   List3.ListIndex = (-1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Label1 = ""
End Sub

Private Sub List1_DblClick()
   TTXX$ = List1.TEXT
   List1.RemoveItem List1.ListIndex
   List2.AddItem TTXX$
End Sub

Private Sub List2_DblClick()
   TTXX$ = List2.TEXT
   List2.RemoveItem List2.ListIndex
   List1.AddItem TTXX$
End Sub

Private Sub MESAN_DblClick()
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
      MESAN = MESASEL$
      MESAN.Refresh
   End If
End Sub

Private Sub Option3_Click(Index As Integer)
   For KI% = 0 To 3
     If Option3(KI%).Value = True Then
       Label1 = Option3(Index).Caption
     End If
   Next KI%
End Sub
