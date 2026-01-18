VERSION 5.00
Begin VB.Form OPORFABRI 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración y Formulario de Orden de Fabricación"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10605
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
   ScaleHeight     =   7290
   ScaleWidth      =   10605
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command9 
      Caption         =   "Más ..."
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
      Left            =   5775
      Picture         =   "OPORFABRI.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Otras Opciones de Configuración"
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Print"
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
      Left            =   6720
      Picture         =   "OPORFABRI.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Imprime Formulario de Configuración"
      Top             =   210
      Width           =   855
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Orden de Fabricación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6000
      Left            =   210
      TabIndex        =   3
      Top             =   1155
      Width           =   10200
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "M.Prima Unica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   8190
         TabIndex        =   41
         Top             =   2730
         Width           =   1815
         Begin VB.OptionButton Option5 
            BackColor       =   &H00C0FFFF&
            Caption         =   "V.Materiales"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   43
            Top             =   315
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Form.Ppal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   42
            Top             =   630
            Width           =   1380
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Impresión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   2835
         TabIndex        =   39
         Top             =   1575
         Width           =   2130
         Begin VB.CheckBox Check14 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Suprime Impre-  sión (Oculta)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   315
            TabIndex        =   40
            Top             =   315
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Documentos Anexos (Requiere Sistema de Calidad)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   210
         TabIndex        =   33
         Top             =   2730
         Width           =   7785
         Begin VB.CheckBox Check12 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Hoja de Inspección"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   6090
            TabIndex        =   37
            Top             =   315
            Width           =   1485
         End
         Begin VB.CheckBox Check11 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Hojas de Procesos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   4305
            TabIndex        =   36
            Top             =   315
            Width           =   1485
         End
         Begin VB.CheckBox Check10 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Especificaciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1995
            TabIndex        =   35
            Top             =   315
            Width           =   1905
         End
         Begin VB.CheckBox Check9 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Plano"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   525
            TabIndex        =   34
            Top             =   315
            Width           =   1065
         End
      End
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
         Height          =   880
         Left            =   9135
         Picture         =   "OPORFABRI.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   525
         Width           =   855
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
         Left            =   5145
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1050
         Width           =   3855
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
         Height          =   360
         Left            =   8820
         TabIndex        =   24
         Top             =   525
         Width           =   175
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
         Left            =   7455
         TabIndex        =   23
         Text            =   " "
         Top             =   525
         Width           =   1380
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fórmula de Consumo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   6195
         TabIndex        =   17
         Top             =   1575
         Width           =   3795
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Multinivel Linealizada"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   2205
            TabIndex        =   19
            Top             =   315
            Width           =   1485
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Lineal (Un Nivel de Explosión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   210
            TabIndex        =   18
            Top             =   315
            Value           =   -1  'True
            Width           =   1800
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Controles Varios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   210
         TabIndex        =   14
         Top             =   3780
         Width           =   9780
         Begin VB.CheckBox Check13 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Exige Declarar Sector Productivo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5565
            TabIndex        =   38
            Top             =   1575
            Width           =   4110
         End
         Begin VB.CheckBox Check8 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Múltiplos de Lotes Económicos (Tandas)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5565
            TabIndex        =   30
            Top             =   945
            Width           =   4110
         End
         Begin VB.CheckBox Check7 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Vale Individual por Operación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5565
            TabIndex        =   29
            Top             =   630
            Width           =   4110
         End
         Begin VB.CheckBox Check6 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Fracciona por Lotes Máximos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5565
            TabIndex        =   28
            Top             =   315
            Width           =   4110
         End
         Begin VB.CheckBox Check5 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Presenta Componentes Consumidos al Cierre"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   525
            TabIndex        =   22
            Top             =   1575
            Width           =   4110
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Presenta Componentes al Lanzamiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   525
            TabIndex        =   21
            Top             =   1260
            Width           =   4110
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Verifica Componentes al Lanzamiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   525
            TabIndex        =   20
            Top             =   945
            Width           =   4110
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Observaciones - Leyendas - Comentarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   525
            TabIndex        =   16
            Top             =   630
            Width           =   4110
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Generar O/F's por Subconjuntos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   525
            TabIndex        =   15
            Top             =   315
            Width           =   4110
         End
      End
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
         Left            =   2415
         TabIndex        =   13
         Text            =   " "
         Top             =   525
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
         Height          =   360
         Left            =   3780
         TabIndex        =   12
         Top             =   525
         Width           =   175
      End
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
         Left            =   210
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1050
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
         Height          =   880
         Left            =   4095
         Picture         =   "OPORFABRI.frx":0C7E
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   525
         Width           =   855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   210
         TabIndex        =   6
         Top             =   1575
         Width           =   2450
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Manual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   630
            TabIndex        =   8
            Top             =   315
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Automatica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   630
            TabIndex        =   7
            Top             =   630
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   5145
         TabIndex        =   4
         Top             =   1575
         Width           =   855
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
            Left            =   190
            TabIndex        =   5
            Text            =   "OF"
            Top             =   420
            Width           =   450
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario Auxiliar:"
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
         Left            =   5040
         TabIndex        =   27
         Top             =   630
         Width           =   2325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario Principal:"
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
         TabIndex        =   9
         Top             =   630
         Width           =   2325
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
      Left            =   7665
      Picture         =   "OPORFABRI.frx":0F88
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
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
      Left            =   8610
      Picture         =   "OPORFABRI.frx":1292
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   855
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
      Height          =   855
      Left            =   9555
      Picture         =   "OPORFABRI.frx":159C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   210
      Width           =   855
   End
End
Attribute VB_Name = "OPORFABRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "FORMORFA", Text1.TEXT)
    Call GRACONTROL("", "FORMORFB", Text3.TEXT)
    Call GRACONTROL("", "CODMORFA", Text5.TEXT)
    '
    MONUME$ = "AUTO": If Option1.Value = True Then MONUME$ = "MANUAL"
    Call GRACONTROL("", "NUMEORFA", MONUME$)
    '
    MONUME$ = "LINEAL": FORMUCON% = 1
    If Option4.Value = True Then
       MONUME$ = "MULTILIN"
       FORMUCON% = 2
    End If
    Call GRACONTROL("", "FORMUCON", MONUME$)
    '
    MONUME$ = "NO"
    If Check1.Value = 1 Then
       MONUME$ = "SI"
       If FORMUCON% = 2 Then
          Call MENSERR(24, "Opciones no Compatibles")
          Check1.Value = 0
          Exit Sub
       End If
    End If
    Call GRACONTROL("", "GESUBOR", MONUME$)
    '
    MONUME$ = "NO": If Check2.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("", "OBSERVOF", MONUME$)
    '
    MONUME$ = "NO": CHECOMPO% = (-1)
    If Check3.Value = 1 Then
        MONUME$ = "SI"
        CHECOMPO% = 1
    End If
    Call GRACONTROL("", "CHECOMPO", MONUME$)
    '
    MONUME$ = "NO": PRECOLAN% = (-1)
    If Check4.Value = 1 Then
        MONUME$ = "SI"
        PRECOLAN% = 1
    End If
    Call GRACONTROL("", "PRECOLAN", MONUME$)
    '
    MONUME$ = "NO": PRECOCIE% = (-1)
    If Check5.Value = 1 Then
        MONUME$ = "SI"
        PRECOCIE% = 1
    End If
    Call GRACONTROL("", "PRECOCIE", MONUME$)
    '
    MONUME$ = "NO": FRALOMAX% = (-1)
    If Check6.Value = 1 Then
        MONUME$ = "SI"
        FRALOMAX% = 1
    End If
    Call GRACONTROL("", "FRALOMAX", MONUME$)
    '
    MONUME$ = "NO": VALINDOP% = (-1)
    If Check7.Value = 1 Then
        MONUME$ = "SI"
        VALINDOP% = 1
    End If
    Call GRACONTROL("", "VALINDOP", MONUME$)
    '
    MONUME$ = "NO": VERITAND% = (-1)
    If Check8.Value = 1 Then
        MONUME$ = "SI"
        VERITAND% = 1
    End If
    Call GRACONTROL("", "VERITAND", MONUME$)
    '
    MONUME$ = "NO": EXISEPRO% = (-1)
    If Check13.Value = 1 Then
        MONUME$ = "SI"
        EXISEPRO% = 1
    End If
    Call GRACONTROL("", "EXISEPRO", MONUME$)
    '
    MONUME$ = "NO"
    If Check9.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ORDEFABR", "PRIPLANO", MONUME$)
    '
    MONUME$ = "NO"
    If Check10.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ORDEFABR", "PRIESPEC", MONUME$)
    '
    MONUME$ = "NO"
    If Check11.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ORDEFABR", "PRIHOPRO", MONUME$)
    '
    MONUME$ = "NO"
    If Check12.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ORDEFABR", "PRIHOINS", MONUME$)
    '
    MONUME$ = "NO"
    If Check14.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL$("ORDEFABR", "SUPRINT", MONUME$)
    '
    MONUME$ = "VALEMAT": If Option6.Value = True Then MONUME$ = "INCLUPRI"
    Call GRACONTROL$("ORFABRI", "MPRIUNIQ", MONUME$)
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
'

Private Sub Command4_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command5_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command6_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text3.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command8_Click()
   PRINTFORM
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   OPVALEMAT.Show 1
   Command9.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Text1.TEXT = Control$("", "FORMORFA")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text3.TEXT = Control$("", "FORMORFB")
    Text4.TEXT = ECOARCH$("INDIFRM", Text3.TEXT)
    '
    If Control$("", "NUMEORFA") = "MANUAL" Then Option1.Value = True
    If Control$("", "GESUBOR") = "SI" Then Check1.Value = 1
    If Control$("", "OBSERVOF") = "SI" Then Check2.Value = 1
    If Control$("", "CHECOMPO") = "SI" Then Check3.Value = 1
    If Control$("", "PRECOLAN") = "SI" Then Check4.Value = 1
    If Control$("", "PRECOCIE") = "SI" Then Check5.Value = 1
    If Control$("", "FRALOMAX") = "SI" Then Check6.Value = 1
    If Control$("", "VALINDOP") = "SI" Then Check7.Value = 1
    If Control$("", "VERITAND") = "SI" Then Check8.Value = 1
    If Control$("", "EXISEPRO") = "SI" Then Check13.Value = 1
    If Control$("ORDEFABR", "PRIPLANO") = "SI" Then Check9.Value = 1
    If Control$("ORDEFABR", "PRIESPEC") = "SI" Then Check10.Value = 1
    If Control$("ORDEFABR", "PRIHOPRO") = "SI" Then Check11.Value = 1
    If Control$("ORDEFABR", "PRIHOINS") = "SI" Then Check12.Value = 1
    If Control$("", "FORMUCON") = "MULTILIN" Then Option4.Value = True
    If Control$("ORDEFABR", "SUPRINT") = "SI" Then Check14.Value = 1
    If Control$("ORFABRI", "MPRIUNIQ") = "INCLUPRI" Then Option6.Value = 1
    '
    CMO$ = TRIM$(Control$("", "CODMORFA")): If CMO$ <> "" Then Text5.TEXT = CMO$
    '
    EPAC$ = UCase$(EMPREAC$)
    If InStr(EPAC$, "NEUMANN") > 0 Then
       Command9.Visible = True
    End If
    '
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text4_CLICK()
    Text2.SetFocus
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
