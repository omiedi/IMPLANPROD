VERSION 5.00
Begin VB.Form OPREMICOMO 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración y Formulario de Remito de Expedición"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10770
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
   ScaleHeight     =   5700
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Nota de Recepción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3390
      Index           =   1
      Left            =   5460
      TabIndex        =   28
      Top             =   2205
      Width           =   5160
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0FF&
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
         Height          =   750
         Index           =   1
         Left            =   4095
         TabIndex        =   41
         Top             =   1470
         Width           =   855
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "NR"
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
            Index           =   1
            Left            =   210
            TabIndex        =   45
            Top             =   290
            Width           =   435
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
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
         Height          =   750
         Index           =   1
         Left            =   210
         TabIndex        =   38
         Top             =   1470
         Width           =   3750
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Auto-Numeración"
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
            Index           =   1
            Left            =   1890
            TabIndex        =   40
            Top             =   315
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Pre-Impresa"
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
            Index           =   1
            Left            =   315
            TabIndex        =   39
            Top             =   315
            Value           =   -1  'True
            Width           =   1380
         End
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
         Index           =   1
         Left            =   4095
         Picture         =   "OPREMICOMO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   420
         Width           =   855
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
         Index           =   1
         Left            =   210
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   " "
         Top             =   945
         Width           =   3750
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
         Index           =   1
         Left            =   3780
         TabIndex        =   35
         Top             =   420
         Width           =   175
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
         Index           =   1
         Left            =   2415
         TabIndex        =   34
         Text            =   " "
         Top             =   420
         Width           =   1380
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Papel"
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
         Index           =   1
         Left            =   210
         TabIndex        =   31
         Top             =   2310
         Width           =   3330
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Formulario Continuo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   315
            TabIndex        =   33
            Top             =   315
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Hoja Suelta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   1890
            TabIndex        =   32
            Top             =   315
            Width           =   1065
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Prefijo"
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
         Index           =   1
         Left            =   3675
         TabIndex        =   29
         Top             =   2310
         Width           =   1275
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
            Index           =   1
            Left            =   210
            TabIndex        =   30
            TabStop         =   0   'False
            Text            =   "0001"
            Top             =   315
            Width           =   870
         End
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   0
         TabIndex        =   42
         Top             =   525
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
      Left            =   7875
      Picture         =   "OPREMICOMO.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   315
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
      Left            =   8820
      Picture         =   "OPREMICOMO.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   315
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
      Left            =   9765
      Picture         =   "OPREMICOMO.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame10"
      Height          =   1170
      Left            =   6930
      TabIndex        =   20
      Top             =   105
      Width           =   3690
      Begin VB.CommandButton Command6 
         Caption         =   "More"
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
         Left            =   0
         Picture         =   "OPREMICOMO.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Opciones Complementarias"
         Top             =   210
         Width           =   855
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame12"
      Height          =   960
      Left            =   10730
      TabIndex        =   19
      Top             =   945
      Width           =   3900
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Opciones Avanzadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   210
      TabIndex        =   18
      Top             =   3675
      Width           =   5160
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Talonario en Uso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   210
         TabIndex        =   24
         Top             =   945
         Width           =   4740
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
            Left            =   4370
            TabIndex        =   26
            Top             =   275
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
            Left            =   3215
            TabIndex        =   25
            Text            =   " "
            Top             =   275
            Width           =   1185
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Nueva Numeración:"
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
            Left            =   735
            TabIndex        =   27
            Top             =   315
            Width           =   2325
         End
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Reemplazo por Código Original"
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
         Left            =   3150
         TabIndex        =   23
         Top             =   315
         Width           =   1725
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Verificar Existencia de Piezas a Despachar"
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
         Left            =   420
         TabIndex        =   22
         Top             =   315
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Remito de Expedición"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3390
      Index           =   0
      Left            =   210
      TabIndex        =   3
      Top             =   210
      Width           =   5160
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Prefijo"
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
         Index           =   0
         Left            =   3675
         TabIndex        =   16
         Top             =   2310
         Width           =   1275
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
            Index           =   0
            Left            =   210
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "0001"
            Top             =   315
            Width           =   870
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Papel"
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
         Index           =   0
         Left            =   210
         TabIndex        =   13
         Top             =   2310
         Width           =   3330
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Hoja Suelta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   1890
            TabIndex        =   15
            Top             =   315
            Width           =   1065
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Formulario Continuo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   315
            TabIndex        =   14
            Top             =   315
            Value           =   -1  'True
            Width           =   1275
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
         Index           =   0
         Left            =   2415
         TabIndex        =   12
         Text            =   " "
         Top             =   420
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
         Index           =   0
         Left            =   3780
         TabIndex        =   11
         Top             =   420
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
         Index           =   0
         Left            =   210
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   " "
         Top             =   945
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
         Index           =   0
         Left            =   4095
         Picture         =   "OPREMICOMO.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   420
         Width           =   855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
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
         Height          =   750
         Index           =   0
         Left            =   210
         TabIndex        =   5
         Top             =   1470
         Width           =   3750
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Pre-Impresa"
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
            Index           =   0
            Left            =   315
            TabIndex        =   7
            Top             =   315
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Auto-Numeración"
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
            Index           =   0
            Left            =   1890
            TabIndex        =   6
            Top             =   315
            Width           =   1800
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0FF&
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
         Height          =   750
         Index           =   0
         Left            =   4095
         TabIndex        =   4
         Top             =   1470
         Width           =   855
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "RT"
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
            Index           =   0
            Left            =   210
            TabIndex        =   44
            Top             =   290
            Width           =   435
         End
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   0
         TabIndex        =   8
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame13"
      Height          =   1275
      Left            =   4410
      TabIndex        =   21
      Top             =   105
      Width           =   1035
   End
   Begin VB.Shape Shape1 
      Height          =   540
      Left            =   5440
      Top             =   630
      Width           =   1470
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   5210
      Picture         =   "OPREMICOMO.frx":107C
      Top             =   640
      Width           =   2010
   End
End
Attribute VB_Name = "OPREMICOMO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command2_Click()
    '
    FENUE$ = TRIM$(Text3.TEXT)
    If FENUE$ = "" Then
       FENUE$ = "01/01/80"
    End If
    If FECHANUM(FENUE$) < 1 Then
       Call MENSERR(24, "Fecha no Válida")
       Text3.SetFocus
       Exit Sub
    End If
    '
    VERISTOK$ = "NOVERIF"
    If Check1.Value > 0 Then
        VERISTOK$ = "VERISTOK"
    End If
    Call GRACONTROL("REMSTAN", "VERISTOK", VERISTOK$)
    '
    MONUME$ = "NO"
    If Check2.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("REMSTAN", "REPLAORI", MONUME$)
    '
    Call GRACONTROL("", "FENUREFA", FENUE$)
    '
    For KKI% = 0 To 1
      IKS$ = TRIM$(Str$(KKI%))
      Call GRACONTROL("", "FOREMIT" + IKS$, Text1(KKI%).TEXT)
      '
      MONUME$ = "AUTONUME": If Option1(KKI%).Value = True Then MONUME$ = "PREIMPRE"
      Call GRACONTROL("", "NUREMIT" + IKS$, MONUME$)
      '
      TIPAPEL$ = "FORMUCON": If Option4(KKI%).Value = True Then TIPAPEL$ = "HOJASUEL"
      Call GRACONTROL("", "TIPAREM" + IKS$, TIPAPEL$)
      '
      PREFIREM$ = TRIM$(Text4(KKI%).TEXT): If PREFIREM$ = "" Then PREFIREM$ = "0001"
      If Len(PREFIREM$) <> 4 Then
        Call COMUNI("Prefijo no Válido")
        Text4(KKI%).SetFocus
        Exit Sub
      End If
      For U& = 1 To 4
        UI% = Asc(Mid$(PREFIREM$, U&))
        If UI% < 48 Or UI% > 57 Then
           Call COMUNI("Prefijo no Válido")
           Text4(KKI%).SetFocus
           Exit Sub
        End If
      Next U&
      Call GRACONTROL("", "PREFIRE" + IKS$, PREFIREM$)
      '
    Next KKI%
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command3_Click(Index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1(Index).TEXT = VDEVU$
        Text2(Index).TEXT = VALACT2$("INDIFRM")
    End If
End Sub
'

Private Sub Command4_Click()
    Call SELECHO("SELFECHA")
    FENUE$ = TRIM$(VALACT1$("SELFECHA"))
    If FENUE$ <> "" Then
        Text3.TEXT = FENUE$
    End If
End Sub

Private Sub Command5_Click(Index As Integer)
    Call CONECRUN("FLFORMS/" + TRIM$(Text1(Index).TEXT), "Formularios de Impresion")
    If EXISTE%("FLFORMS.EXE") < 1 Then
        Call FINAL("")
    End If
End Sub

Private Sub Command6_Click()
  Call OPNOIN("")
End Sub

Private Sub Form_Load()
    '
    For KKI% = 0 To 1
      IKS$ = TRIM$(Str$(KKI%))
      Text1(KKI%).TEXT = CONTROL$("", "FOREMIT" + IKS$)
      Text2(KKI%).TEXT = ECOARCH$("INDIFRM", Text1(KKI%).TEXT)
      '
      If CONTROL$("", "NUREMIT" + IKS$) = "AUTONUME" Then Option2(KKI%).Value = True
      If CONTROL$("", "TIPAREM" + IKS$) = "HOJASUEL" Then Option4(KKI%).Value = True
      '
       PREFIREM$ = TRIM$(CONTROL$("", "PREFIRE" + IKS$)): If PREFIREM$ = "" Then PREFIREM$ = "0001"
       Text4(KKI%).TEXT = PREFIREM$
    Next KKI%
    Text3.TEXT = CONTROL$("", "FENUREFA")
    '
    If CONTROL$("REMSTAN", "VERISTOK") = "VERISTOK" Then
       Check1.Value = 1
    End If
    '
    If CONTROL$("REMSTAN", "REPLAORI") = "SI" Then
       Check2.Value = 1
    End If
End Sub

Private Sub Text1_DblClick(Index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1(Index).TEXT = VDEVU$
        Text2(Index).TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text2_CLICK(Index As Integer)
    Text1(Index).SetFocus
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("SELFECHA")
    FENUE$ = TRIM$(VALACT1$("SELFECHA"))
    If FENUE$ <> "" Then
        Text3.TEXT = FENUE$
    End If
End Sub

