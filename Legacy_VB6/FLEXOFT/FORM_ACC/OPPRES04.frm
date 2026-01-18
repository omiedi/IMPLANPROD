VERSION 5.00
Begin VB.Form OPPRES04 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Funcionamiento - Presupuestos"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   6270
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
   ScaleHeight     =   6255
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Height          =   7500
      Left            =   5355
      ScaleHeight     =   7440
      ScaleWidth      =   1110
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
      Begin VB.CommandButton Command11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   120
         Picture         =   "OPPRES04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   5400
         Width           =   650
      End
      Begin VB.CommandButton Command10 
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
         Left            =   120
         Picture         =   "OPPRES04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Guardar "
         Top             =   7020
         Width           =   640
      End
      Begin VB.Image Image3 
         Height          =   7500
         Left            =   0
         Picture         =   "OPPRES04.frx":0614
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Factura Proforma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   210
      TabIndex        =   38
      Top             =   2325
      Width           =   4950
      Begin VB.CommandButton Command8 
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
         Height          =   720
         Left            =   4080
         Picture         =   "OPPRES04.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   330
         Width           =   645
      End
      Begin VB.TextBox Text8 
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
         Left            =   195
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   " "
         Top             =   765
         Width           =   3750
      End
      Begin VB.CommandButton Command9 
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
         Left            =   3735
         TabIndex        =   40
         Top             =   345
         Width           =   175
      End
      Begin VB.TextBox Text7 
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
         Left            =   2370
         TabIndex        =   39
         Text            =   " "
         Top             =   345
         Width           =   1380
      End
      Begin VB.Label Label4 
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
         Index           =   2
         Left            =   0
         TabIndex        =   43
         Top             =   450
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Controles Operativos - Campos Mandatorios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   210
      TabIndex        =   12
      Top             =   6570
      Width           =   4950
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Plazo de Garantía"
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
         Index           =   8
         Left            =   2900
         TabIndex        =   26
         Top             =   1365
         Width           =   1920
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Validez de Oferta"
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
         Index           =   7
         Left            =   2900
         TabIndex        =   25
         Top             =   1050
         Width           =   1920
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Lugar de Entrega"
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
         Index           =   6
         Left            =   2900
         TabIndex        =   24
         Top             =   735
         Width           =   1920
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Plazo de Entrega"
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
         Index           =   5
         Left            =   2900
         TabIndex        =   23
         Top             =   420
         Width           =   1920
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Vendedor"
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
         Index           =   4
         Left            =   350
         TabIndex        =   22
         Top             =   1680
         Width           =   2250
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Condición de Pago"
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
         Index           =   3
         Left            =   350
         TabIndex        =   21
         Top             =   1365
         Width           =   4005
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Lista de Precios"
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
         Index           =   2
         Left            =   350
         TabIndex        =   20
         Top             =   1050
         Width           =   4005
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Referencia"
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
         Left            =   350
         TabIndex        =   19
         Top             =   735
         Width           =   4005
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Atención Sr./Sra./Srta"
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
         Left            =   350
         TabIndex        =   18
         Top             =   420
         Width           =   4005
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Formulario Presupuesto - Cotización"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   210
      TabIndex        =   6
      Top             =   90
      Width           =   4950
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
         Height          =   285
         Index           =   1
         Left            =   2415
         TabIndex        =   30
         Text            =   " "
         Top             =   1260
         Width           =   1380
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
         Height          =   285
         Index           =   1
         Left            =   3780
         TabIndex        =   29
         Top             =   1260
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
         Height          =   285
         Index           =   1
         Left            =   210
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1680
         Width           =   3750
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
         Height          =   720
         Index           =   1
         Left            =   4095
         Picture         =   "OPPRES04.frx":2CB6
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1240
         Width           =   645
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
         Height          =   720
         Index           =   0
         Left            =   4095
         Picture         =   "OPPRES04.frx":2FC0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   300
         Width           =   645
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
         Height          =   285
         Index           =   0
         Left            =   210
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
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
         Height          =   285
         Index           =   0
         Left            =   3780
         TabIndex        =   8
         Top             =   315
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
         Height          =   285
         Index           =   0
         Left            =   2415
         TabIndex        =   7
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Lista de  Repuestos:"
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
         TabIndex        =   31
         Top             =   1365
         Width           =   2325
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item Unico:"
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
         TabIndex        =   11
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Anexos - Documentos Graficos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2520
      Left            =   210
      TabIndex        =   0
      Top             =   3690
      Width           =   4950
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Control de Impresion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   210
         TabIndex        =   13
         Top             =   1365
         Width           =   4530
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Fotografía - 2"
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
            Index           =   3
            Left            =   2520
            TabIndex        =   17
            Top             =   630
            Width           =   1905
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Fotografía - 1"
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
            Index           =   2
            Left            =   2520
            TabIndex        =   16
            Top             =   315
            Width           =   1905
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Croquis"
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
            Left            =   525
            TabIndex        =   15
            Top             =   630
            Width           =   1905
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Plano de Pieza"
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
            Left            =   525
            TabIndex        =   14
            Top             =   315
            Width           =   1905
         End
      End
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
         Height          =   285
         Left            =   2415
         TabIndex        =   5
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
         Height          =   285
         Left            =   3780
         TabIndex        =   4
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text6 
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
         Left            =   210
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
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
         Height          =   720
         Left            =   4095
         Picture         =   "OPPRES04.frx":32CA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   315
         Width           =   645
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
         Left            =   0
         TabIndex        =   1
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808000&
      Height          =   7500
      Left            =   5355
      ScaleHeight     =   7440
      ScaleWidth      =   1110
      TabIndex        =   32
      Top             =   0
      Width           =   1170
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   36
         Top             =   5385
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "OPPRES04.frx":35D4
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   315
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   0
         Left            =   105
         Picture         =   "OPPRES04.frx":371E
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   955
         Width           =   650
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "OPPRES04.frx":3A28
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   4545
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "OPPRES04.frx":3D32
         Top             =   5655
         Width           =   1515
      End
   End
   Begin VB.Menu ar 
      Caption         =   "Archivo"
      Begin VB.Menu guardar 
         Caption         =   "Guardar"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "OPPRES04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command11_Click()
    Call Command2_Click
End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "FORPRESU", Text1(0).TEXT)
    Call GRACONTROL("", "FORPRERE", Text1(1).TEXT)
    
    Call GRACONTROL("", "FORGRAFI", Text5.TEXT)
    Call GRACONTROL("", "FORFAPRO", Text7.TEXT)
    '
    For KKI% = 0 To 3
      MONUME$ = "NO": If Check1(KKI%).Value > 0 Then MONUME$ = "SI"
      Call GRACONTROL("PRESUPUE", "PRIGRA-" + TRIM$(Str$(KKI%)), MONUME$)
    Next KKI%
    '
    For KKI% = 0 To 8
      MONUME$ = "NO": If Check2(KKI%).Value > 0 Then MONUME$ = "SI"
      Call GRACONTROL("PRESUPUE", "CTRLX-" + TRIM$(Str$(KKI%)), MONUME$)
    Next KKI%
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text5.TEXT = VDEVU$
        Text6.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
'

Private Sub Command4_Click(Index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1(Index).TEXT = VDEVU$
        Text2(Index).TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command5_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text5.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command6_Click(Index As Integer)
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1(Index).TEXT), "Formularios de Impresion")
End Sub

Private Sub Command8_Click()
  ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
  Call CONECRUN(ACONEC$ + "/" + TRIM$(Text7.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command9_Click()
   Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Text8.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Form_Load()
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture2.Visible = True
      Me.Frame1.BackColor = &H8000000F
      Me.Frame2.BackColor = &H8000000F
      Me.Frame3.BackColor = &H8000000F
      Me.Frame4.BackColor = &H8000000F
      Me.Frame5.BackColor = &H8000000F
      For I% = 0 To Me.Check1.UBound
         Me.Check1(I%).BackColor = &H8000000F
      Next I%
      For I% = 0 To Me.Check2.UBound
         Me.Check2(I%).BackColor = &H8000000F
      Next I%
      Me.BackColor = &H8000000F
  End If

    Text1(0).TEXT = CONTROL$("", "FORPRESU")
    Text2(0).TEXT = ECOARCH$("INDIFRM", Text1(0).TEXT)
    Text1(1).TEXT = CONTROL$("", "FORPRERE")
    Text2(1).TEXT = ECOARCH$("INDIFRM", Text1(1).TEXT)
    Text5.TEXT = CONTROL$("", "FORGRAFI")
    Text6.TEXT = ECOARCH$("INDIFRM", Text5.TEXT)
    Text7.TEXT = TRIM$(CONTROL$("", "FORFAPRO"))
    Text8.TEXT = ECOARCH$("INDIFRM", Text7)
    '
    For KKI% = 0 To 3
      Check1(KKI%).Value = 0
      If CONTROL$("PRESUPUE", "PRIGRA-" + TRIM$(Str$(KKI%))) = "SI" Then
        Check1(KKI%).Value = 1
      End If
    Next KKI%
    '
    For KKI% = 0 To 8
      Check2(KKI%).Value = 0
      If CONTROL$("PRESUPUE", "CTRLX-" + TRIM$(Str$(KKI%))) = "SI" Then
        Check2(KKI%).Value = 1
      End If
    Next KKI%
    '
End Sub

Private Sub guardar_Click()
     Call Command2_Click
End Sub

Private Sub salir_Click()
   Call Command1_Click
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

Private Sub Text6_Click()
    Text5.SetFocus
End Sub

Private Sub Text5_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text5.TEXT = VDEVU$
        Text6.TEXT = VALACT2$("INDIFRM")
    End If
End Sub


Private Sub Text7_DblClick()
   Call Command9_Click
End Sub
