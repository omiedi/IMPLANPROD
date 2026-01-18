VERSION 5.00
Begin VB.Form FORVENTA 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11010
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox LIREGSEL 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   4620
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5355
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2430
      Left            =   7350
      Max             =   1
      Min             =   1
      TabIndex        =   16
      Top             =   1785
      Value           =   1
      Width           =   240
   End
   Begin VB.Frame FRAMEBOT2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2115
      Left            =   9135
      TabIndex        =   10
      Top             =   3465
      Width           =   960
      Begin VB.CommandButton BOTEDIT 
         Caption         =   "Edit"
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
         Left            =   0
         TabIndex        =   15
         ToolTipText     =   "Editar Registro para Actualización"
         Top             =   945
         Width           =   750
      End
      Begin VB.CommandButton BOTAGRE 
         Caption         =   "Add"
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
         Left            =   0
         TabIndex        =   14
         ToolTipText     =   "Agregar Registro al Final"
         Top             =   0
         Width           =   750
      End
      Begin VB.CommandButton BOTINSE 
         Caption         =   "Ins."
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
         Left            =   0
         TabIndex        =   13
         ToolTipText     =   "Intercalar Registro Antes del Activo"
         Top             =   315
         Width           =   750
      End
      Begin VB.CommandButton BOTBUSCA 
         Caption         =   "Look"
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
         Left            =   0
         TabIndex        =   12
         ToolTipText     =   "Buscar Texto por Campo"
         Top             =   1260
         Width           =   750
      End
      Begin VB.CommandButton BOTERASE 
         Caption         =   "Del."
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
         Left            =   0
         TabIndex        =   11
         ToolTipText     =   "Borrar Registro Activo"
         Top             =   630
         Width           =   750
      End
   End
   Begin VB.Frame FRAMEBOT1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2115
      Left            =   9135
      TabIndex        =   7
      Top             =   105
      Width           =   855
      Begin VB.CommandButton BOTEXIT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   0
         Picture         =   "FORVENTABU.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cancelar la Aplicación"
         Top             =   0
         Width           =   750
      End
      Begin VB.CommandButton BOTREC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   0
         Picture         =   "FORVENTABU.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Grabar - Guardar Modificaciones"
         Top             =   570
         Width           =   750
      End
   End
   Begin VB.Frame FRAMETOT 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   645
      Left            =   0
      TabIndex        =   5
      Top             =   4200
      Visible         =   0   'False
      Width           =   7575
      Begin VB.Label TOTALI 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   105
         TabIndex        =   6
         Top             =   210
         Width           =   120
      End
   End
   Begin VB.Frame FRAMENCA 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   645
      Left            =   0
      TabIndex        =   3
      Top             =   1155
      Width           =   7575
      Begin VB.Label ENCACOL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   105
         TabIndex        =   4
         Top             =   210
         Width           =   720
      End
   End
   Begin VB.ListBox LITAB 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   0
      TabIndex        =   2
      Top             =   1785
      Width           =   7575
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   105
      Top             =   6195
   End
   Begin VB.ListBox SORTLIST 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   735
      Sorted          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5355
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3465
      TabIndex        =   1
      Top             =   6405
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "FORVENTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTAGRE_Click()
    Call ADDLISVEN
End Sub

Private Sub BOTBUSCA_Click()
    Call BULISVEN
End Sub

Private Sub BOTEDIT_Click()
    Call EDLISVEN
End Sub

Private Sub BOTERASE_Click()
    Call BORLISVEN
    MODIFI% = 1
End Sub

Private Sub BOTHELP_Click()
    Call HELPONLINE("WINTABU")
End Sub

Private Sub BOTINSE_Click()
    Call INSLISVEN
End Sub

Private Sub BOTREC_Click()
    '
    MODIFI% = 0
    VTB% = 0
    Call CIERRARCH("*.*")
    FORVENTA.Hide
    Call FRESHALL
    '
End Sub



Private Sub BOTEXIT_Click()
    '
If App.EXEName = "PRUVENTB" Then
MsgBox "WIDTH= " & Width '
'& "   " & OFS%
'& "   " & (RGF& - RGI&)
End If
    If MODIFI% > 0 Then
        RESPU = MsgBox("Abandona sin Guardar los Cambios ?", 4371, "Operacion Critica")
        If RESPU <> 6 Then
            BOTREC.SetFocus
            Exit Sub
        End If
        MODIFI% = 0
    End If
    '
    VTB% = (-1)
    Call CIERRARCH("*.*")
    FORVENTA.Hide
    Call FRESHALL
    '
End Sub
'

Private Sub Form_Activate()
    '
    On Error Resume Next
    SetFocus
'    Refresh
    On Error GoTo 0
    If MODEDI% = 1 Then
        If YAPRE% = 0 Then
            On Error Resume Next
            For U& = 0 To Controls.Count - 1
                Controls(U&).Refresh
            Next U&
            On Error GoTo 0
            Call ADDLISVEN
        End If
    End If
    '
End Sub
'


Private Sub Form_Load()
    '
    YAFIL% = 0
    Enabled = True
    '
End Sub

Private Sub List1_Click()

End Sub

Private Sub LITAB_DblClick()
    Call EDLISVEN
End Sub
