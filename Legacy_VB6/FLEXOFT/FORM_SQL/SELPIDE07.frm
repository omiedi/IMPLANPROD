VERSION 5.00
Begin VB.Form SELPIDE07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección Deposito Origen para Lista de Picking"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   285
      Left            =   3300
      MaxLength       =   3
      TabIndex        =   10
      Top             =   1080
      Width           =   675
   End
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
      Height          =   285
      Left            =   3990
      TabIndex        =   7
      Top             =   240
      Width           =   175
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004040&
      Height          =   6000
      Left            =   4305
      ScaleHeight     =   5940
      ScaleWidth      =   975
      TabIndex        =   4
      Top             =   0
      Width           =   1035
      Begin VB.CommandButton Command29 
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
         Height          =   630
         Left            =   90
         Picture         =   "SELPIDE07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Aprueba la Selección"
         Top             =   1170
         Width           =   650
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
         Height          =   630
         Left            =   105
         Picture         =   "SELPIDE07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   75
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "SELPIDE07.frx":0454
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   735
         Width           =   1515
      End
   End
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
      Height          =   285
      Left            =   3975
      TabIndex        =   1
      Top             =   1095
      Width           =   175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3315
      TabIndex        =   9
      Top             =   225
      Width           =   675
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   1410
      Width           =   4035
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Deposito Origen:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1545
      TabIndex        =   3
      Top             =   1125
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   555
      Width           =   4035
   End
   Begin VB.Label Label44 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de Picking:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1755
      TabIndex        =   0
      Top             =   270
      Width           =   1485
   End
End
Attribute VB_Name = "SELPIDE07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ListaSel As String
Public DepOrig As String

Private Sub Command1_Click()
    '
    Call CIERRARCH("*.*")
    Unload Me
    '
End Sub

Private Sub Command2_Click()
   
    Call CARGALISEL("INDIPICK", "BALAREQMO", "NULISPICK/F6.0;DESTINO/A16;DELISPICK/A24", "NULISPICK > 0 AND KILO_ASIG > 0", "DISTINCT")
    Call SELECHO("INDIPICK")
    COD$ = VALACT1("INDIPICK")
    If COD$ <> "" Then
       Label4 = COD$
       Label2 = VALACT2("INDIPICK")
    End If
   
End Sub

Private Sub Command29_Click()
    Command29.Enabled = False
    If Label4 = "" Then
       Call MENSERR(24, "Lista de Picking Inexistente o no Válido")
       GoTo 99
    End If
    '
    If Label1 = "" Then
      Call MENSERR(24, "Deposito Origen Invalido")
      GoTo 99
    End If
    '
    DEPOTRANS% = XVALO(CONTROL$("BOLRECEP", "DEPTRANS")) ' DEPOSITO MATERIAL EN TRANSITO
    If XVALO(Text3) = DEPOTRANS% Then
       Call COMUNI("Depósito Seleccionado No Válido\Configurado como Material en Transito ")
       GoTo 99
    End If

    ListaSel = Label4
    DepOrig = Text3
    Call BLANCFORMU
    '
    Command29.Enabled = True
    Hide
    Exit Sub
    '
99  Command29.Enabled = True
End Sub

Private Sub Command5_Click()
  Call SELECHO("TADEPOSI")
  DEPO$ = VALACT1$("TADEPOSI")
    
  If DEPO$ <> "" Then
      Text3 = DEPO$
      Label1 = VALACT2$("TADEPOSI")
  End If
End Sub



Private Sub Label4_DblClick()
   Call Command2_Click
End Sub

Private Sub Text3_DblClick()
   Call Command5_Click
End Sub
Private Sub BLANCFORMU()
    
    Text3 = ""
    Label4 = ""
    Label1 = ""
    Label2 = ""
    
End Sub

Private Sub Form_Load()
  ListaSel = ""
  DepOrig = ""
  Call BLANCFORMU
End Sub

Private Sub Text3_Change()
    Label1 = ""
    If XVALO(Text3) > 0 Then
       Label1 = ECOARCH("TADEPOSI", Chr$(XVALO(Text3)))
    End If
End Sub
