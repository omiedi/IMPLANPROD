VERSION 5.00
Begin VB.Form OPCALS07 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones de Cálculo"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   2415
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   13
      Top             =   105
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00000080&
      Height          =   2400
      Left            =   4560
      ScaleHeight     =   2340
      ScaleWidth      =   870
      TabIndex        =   10
      Top             =   0
      Width           =   930
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
         Height          =   650
         Left            =   90
         Picture         =   "OPCALS07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Inicia Proceso de Cálculo Estadístico"
         Top             =   960
         Width           =   650
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
         Height          =   650
         Left            =   90
         Picture         =   "OPCALS07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -325
         Picture         =   "OPCALS07.frx":0454
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   1950
         Width           =   1515
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFC0&
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
      Top             =   1365
      Width           =   4005
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
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
         Left            =   2520
         TabIndex        =   8
         Top             =   400
         Width           =   600
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
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
         Left            =   1680
         TabIndex        =   7
         Top             =   400
         Width           =   600
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
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
         Left            =   945
         TabIndex        =   6
         Top             =   400
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
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
         Left            =   3255
         TabIndex        =   9
         Top             =   250
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
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
      Top             =   315
      Width           =   4005
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
         Left            =   3600
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
         Width           =   2355
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
Attribute VB_Name = "OPCALS07"
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

Private Sub Form_Load()
   Call APLICACIONSKINS(Me)

   Call CARSELMESA
   MESAN = Left$(REGLEIDO$("SELMESA", 4), 14)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Label1 = ""
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
