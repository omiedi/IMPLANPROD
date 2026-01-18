VERSION 5.00
Begin VB.Form PLANCUEN 
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan General de Cuentas"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11910
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2305
      Left            =   10760
      TabIndex        =   35
      Top             =   120
      Width           =   1040
      Begin VB.CommandButton Command1 
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
         Height          =   610
         Left            =   120
         Picture         =   "PLANCUEN.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Aprueba y Graba el Plan de Cuentas"
         Top             =   920
         Width           =   750
      End
      Begin VB.CommandButton Command9 
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
         Height          =   610
         Left            =   120
         Picture         =   "PLANCUEN.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   300
         Width           =   750
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
         Height          =   610
         Left            =   120
         Picture         =   "PLANCUEN.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1540
         Width           =   750
      End
   End
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   10680
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0E0C0&
      Caption         =   "A-B-M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2920
      Left            =   10760
      TabIndex        =   31
      Top             =   2780
      Width           =   1040
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   610
         Left            =   120
         Picture         =   "PLANCUEN.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Búsqueda de Cuenta por Código o Descripcion"
         Top             =   1540
         Width           =   765
      End
      Begin VB.CommandButton Command6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   610
         Left            =   120
         Picture         =   "PLANCUEN.frx":1AC4
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Imprime Listado de Cuentas Contables "
         Top             =   2160
         Width           =   765
      End
      Begin VB.CommandButton Command24 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   610
         Left            =   125
         Picture         =   "PLANCUEN.frx":1DCE
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Anulación de Cuenta Contable"
         Top             =   920
         Width           =   765
      End
      Begin VB.CommandButton Command21 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   610
         Left            =   125
         Picture         =   "PLANCUEN.frx":20D8
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Alta de Nueva Cuenta Contable"
         Top             =   300
         Width           =   750
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   10760
      ScaleHeight     =   75
      ScaleWidth      =   975
      TabIndex        =   9
      Top             =   2520
      Width           =   1040
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Edición de Cuenta Activa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   240
      TabIndex        =   7
      Top             =   5880
      Width           =   11550
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0C0&
         Caption         =   "Status"
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
         Left            =   6720
         TabIndex        =   27
         Top             =   1200
         Width           =   975
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0E0C0&
            Caption         =   "Baja"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   29
            Top             =   530
            Width           =   735
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0E0C0&
            Caption         =   "Act."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   290
            Width           =   735
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   2520
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   480
         Width           =   4005
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Text            =   "PLANCUEN.frx":2222
         Top             =   1290
         Width           =   6045
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0C0&
         Caption         =   "Imputable"
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
         Left            =   8640
         TabIndex        =   20
         Top             =   390
         Width           =   1695
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0E0C0&
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   22
            Top             =   290
            Width           =   615
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0E0C0&
            Caption         =   "Si"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   21
            Top             =   290
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0C0&
         Caption         =   "Tipo de Cuenta"
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
         Left            =   7800
         TabIndex        =   15
         Top             =   1200
         Width           =   2535
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0C0&
            Caption         =   "Ingresos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   18
            Top             =   290
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0C0&
            Caption         =   "Gastos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   1320
            TabIndex        =   19
            Top             =   530
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0C0&
            Caption         =   "Haber"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   17
            Top             =   530
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0C0&
            Caption         =   "Debe"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   16
            Top             =   290
            Width           =   1215
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   560
         Left            =   8160
         Max             =   4
         TabIndex        =   13
         Top             =   480
         Width           =   220
      End
      Begin VB.CommandButton Command8 
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
         Height          =   390
         Left            =   10630
         Picture         =   "PLANCUEN.frx":2224
         TabIndex        =   4
         ToolTipText     =   "Edición de la Cuenta Activa"
         Top             =   1650
         Width           =   765
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Del"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10630
         Picture         =   "PLANCUEN.frx":252E
         TabIndex        =   3
         ToolTipText     =   "Borra la Cuenta Activa del Plan"
         Top             =   1260
         Width           =   765
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Ins"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10630
         Picture         =   "PLANCUEN.frx":2838
         TabIndex        =   2
         ToolTipText     =   "Intercala Cuenta en el Plan "
         Top             =   870
         Width           =   765
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   480
         TabIndex        =   0
         Top             =   720
         Width           =   1800
      End
      Begin VB.CommandButton Command11 
         Caption         =   "App"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10630
         Picture         =   "PLANCUEN.frx":2B42
         TabIndex        =   1
         ToolTipText     =   "Agrega Cuenta al Final del Plan"
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   7000
         TabIndex        =   41
         Top             =   780
         Width           =   555
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C-I:"
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
         Left            =   6240
         TabIndex        =   40
         Top             =   600
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   10500
         X2              =   10500
         Y1              =   120
         Y2              =   2200
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
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
         Left            =   1140
         TabIndex        =   30
         Top             =   330
         Width           =   1320
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel:"
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
         Left            =   7000
         TabIndex        =   26
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
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
         Left            =   120
         TabIndex        =   24
         Top             =   1050
         Width           =   1920
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   7800
         TabIndex        =   14
         Top             =   480
         Width           =   405
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   960
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   5655
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   10380
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   5550
         Index           =   1
         Left            =   2630
         TabIndex        =   47
         Top             =   -100
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   5550
         Index           =   5
         Left            =   9500
         TabIndex        =   46
         Top             =   -100
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   5550
         Index           =   4
         Left            =   9000
         TabIndex        =   45
         Top             =   -100
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   5550
         Index           =   3
         Left            =   8140
         TabIndex        =   44
         Top             =   -100
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   5550
         Index           =   2
         Left            =   2630
         TabIndex        =   43
         Top             =   -100
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   5550
         Index           =   0
         Left            =   960
         TabIndex        =   42
         Top             =   -100
         Width           =   15
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4980
         ItemData        =   "PLANCUEN.frx":2E4C
         Left            =   0
         List            =   "PLANCUEN.frx":2E4E
         TabIndex        =   6
         Top             =   480
         Width           =   10380
      End
      Begin VB.PictureBox Picture1 
         Height          =   500
         Left            =   0
         ScaleHeight     =   435
         ScaleWidth      =   10305
         TabIndex        =   11
         Top             =   0
         Width           =   10360
         Begin VB.Line Line4 
            X1              =   10340
            X2              =   10320
            Y1              =   360
            Y2              =   0
         End
         Begin VB.Label Label5 
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Nivel   Código        Descripción                                   Tipo   ID  Stat"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   60
            TabIndex        =   12
            Top             =   90
            Width           =   10305
         End
      End
      Begin VB.Line Line6 
         X1              =   10360
         X2              =   10360
         Y1              =   0
         Y2              =   480
      End
   End
   Begin VB.Line Line5 
      X1              =   8275
      X2              =   8280
      Y1              =   4305
      Y2              =   4725
   End
   Begin VB.Line Line3 
      X1              =   8275
      X2              =   8275
      Y1              =   1365
      Y2              =   1680
   End
   Begin VB.Line Line2 
      X1              =   8265
      X2              =   8265
      Y1              =   1365
      Y2              =   1785
   End
End
Attribute VB_Name = "PLANCUEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODOEDIT%, ACTIVLIST$, ILISTACT&
'
Private Sub Command1_Click()
   '
   Call CANCELEDIT
   '
   If CONTROL$("PROPSCG", "PLAGRAL") = "GUEST" Then
      Call MENSERR(24, "Sistema Distribuido - Invitado en Plan de Cuentas\Imposible Grabar Plan de Cuentas.")
      Unload Me
      Exit Sub
   End If
   '
   Call GUARDAPLAN
   Call FINAL("")
   '
End Sub
'
Private Sub Command11_Click()
   '
   Command11.Enabled = False
   Call AGRECUENTA(1)
   Command11.Enabled = True
   '
End Sub

Private Sub Command12_Click()
   '
   Command11.Enabled = False
   Call AGRECUENTA(2)
   Command11.Enabled = True
   '
End Sub

Private Sub Command21_Click()
   Command21.Enabled = False
   '
   Call CANCELEDIT
   '
   If TRIM$(Text11) = "" Then
     Call NUEVACUECON
   End If
   Command21.Enabled = True
End Sub

Sub CANCELEDIT()
   '
   If MODOEDIT% > 0 Then
     '
     List1.AddItem ACTIVLIST$, ILISTACT&
     On Error Resume Next
     List1.ListIndex = ILISTACT&
     On Error GoTo 0
     '
     MODOEDIT% = 0
     ACTIVLIST$ = 0
     ILISTACT& = 0
     Command4.Enabled = True
     Command8.Enabled = True
     Command24.Enabled = True
     '
     Text1.TEXT = ""
     Text11.TEXT = ""
     Label2 = ""
     Label8 = ""
     Option1(0) = True
     Option2(0) = True
     Option3(0) = True
     '
   End If
   '
End Sub

Private Sub Command24_Click()
   '
   Command24.Enabled = False
   If MODOEDIT% > 0 Then
       Call CANCELEDIT
     Else
       U& = List1.ListIndex
       If U& < 0 Then Exit Sub
       If U& > List1.ListCount - 1 Then Exit Sub
       '
       ZZ$ = List1.List(U&)
       List1.RemoveItem U&
       Call REPLA(ZZ$, "Baja", 80, 4)
       List1.AddItem ZZ$, U&
       On Error Resume Next
       List1.ListIndex = U&
       On Error GoTo 0
   End If
   Command24.Enabled = True
   '
End Sub

'
Private Sub Command3_Click()
   '
   Command3.Enabled = False
   '
   Call CANCELEDIT
   '
   Screen.MousePointer = 11
   JJ& = 0
   For U& = 1 To List1.ListCount
     TTXX$ = List1.List(U& - 1)
     TTYY$ = AJUSTI$(Mid$(TTXX$, 9, 12), 12) + AJUSTI$(Mid$(TTXX$, 23, 28), 28)
     JJ& = JJ& + 1
     Call GRAREG("!ECUECOT", TTYY$, JJ&)
   Next U&
   Call SETULTREG("!ECUECOT", JJ&)
   Screen.MousePointer = 1
   Call SELECHO("!ECUECOT")
   '
   ABUS$ = AJUSTI$(VALACT1$("!ECUECOT"), 12)
   For U& = 1 To List1.ListCount
     If Mid$(List1.List(U& - 1), 9, 12) = ABUS$ Then
       List1.ListIndex = U& - 1
       GoTo 99
     End If
   Next U&
   '
99 Command3.Enabled = True
   '
End Sub

Private Sub Command4_Click()
   '
   Command4.Enabled = False
   If MODOEDIT% > 0 Then
       Call CANCELEDIT
     Else
       U& = List1.ListIndex
       If U& < 0 Then Exit Sub
       If U& > List1.ListCount - 1 Then Exit Sub
       '
       ZZ$ = List1.List(U&)
       List1.RemoveItem U&
       Call REPLA(ZZ$, "Baja", 80, 4)
       List1.AddItem ZZ$, U&
       On Error Resume Next
       List1.ListIndex = U&
       On Error GoTo 0
   End If
   Command4.Enabled = True
   '
End Sub

Private Sub Command5_Click()
   '
   JJ& = 0
   FIN& = List1.ListCount
   For U& = 1 To List1.ListCount
      X$ = List1.List(U& - 1)
      JJ& = JJ& + 1
      Call GRAREG("PLANCUE", X$, JJ&)
   Next U&
   '
   Call SETULTREG("PLANCUE", JJ&)
   Call CIERRARCH("PLANCUE")
   Call FINAL("*PLANCUE")
   '
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call CANCELEDIT
   Call FINAL("*LICUECO")
   Command6.Enabled = True
End Sub

'
Private Sub Command8_Click()
   '
   Call EDITACUEN
   '
End Sub

'
Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    Show
    DoEvents
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    Caption = Caption + " - " + TRIM$(EJACT$)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Show
    DoEvents
    '
    Call CARGAPLAN
    '
End Sub
'

Private Sub Form_Unload(Cancel As Integer)
    TIASI% = 0
    Call CIERRACONTA
    Call LIBARCH("ASICONT")
End Sub

Private Sub List1_DblClick()
   '
   Call EDITACUEN
   '
End Sub

Private Sub Command2_Click()
    Call HELPONLINE("PLAGRAL")
End Sub
'

Private Sub Command9_Click()
    '
    Call CANCELEDIT
    '
    Call CIERRARCH("*.*")
    Call FINAL("")
    '
End Sub


Private Sub COMMAND31_CLICK()
   RG1$ = REGSEL$("ECUECOT")
   If Len(RG1$) > 4 Then
      Text11.TEXT = Left$(RG1$, 12)
      Label4.Caption = Mid$(RG1$, 13, 28)
   End If
End Sub
'
Sub RECARPLAN()
    '
    Static YAVERI%
    '
    Screen.MousePointer = 11
    If UCase$(EXTE$) = "OLD" Then YAVERI% = 1
    If YAVERI% < 1 Then
       On Error Resume Next
       If FileLen(LUDAT$ + "PLAGRAL.DAT") >= FileLen(LUDAT$ + "PLAGRAL.OLD") Then YAVERI% = 1
       On Error GoTo 0
    End If
    '
    NXP% = FILEOPEN%(LUDAT$ + "PLAGRAL" + "." + EXTE$, 0, 4)
    Dim IVP As String * 4
    If YAVERI% < 1 Then
      Dim YAPLA%(32767)
      NXQ% = FILEOPEN%(LUDAT$ + "PLAGRAL" + ".OLD", 0, 4)
      Dim IVQ As String * 4
    End If
    '
    List1.Clear
    JJ& = 0
    '
    UCUEPLAN$ = "": UREPLA% = 0
    FIN& = LOF(NXP%) / 4
    For KK% = 1 To LOF(NXP%) / 4
       UU& = KK%
       Call TRACE(20, UU&, FIN&)
       Get #NXP%, KK%, IVP$
       CUE% = CVI(Mid$(IVP$, 1, 2))
       NIVEI% = CVI(Mid$(IVP$, 3, 2))
       If NIVEI% < 0 Then NIVEI% = 0
       If NIVEI% > 4 Then NIVEI% = 4
       '
       If CUE% > 0 Then
         ZZ$ = Space$(64)
         If NIVEI% >= 0 Then
           If NIVEI% <= 4 Then
             NVX$ = Left$(String$(NIVEI%, ".") + TRIM$(Str$(NIVEI%)) + ".....", 5)
             Call REPLA(ZZ$, NVX$, 1, 5)
             Call REPLA(ZZ$, CODCUE$(CUE%), 10, 12)
             Call REPLA(ZZ$, DENOCUE$(CUE%), 26, 32)
             List1.AddItem ZZ$
             UCUEPLAN$ = Mid$(ZZ$, 10, 12)
             If YAVERI% < 1 Then YAPLA%(CUE%) = 1
             UREPLA% = KK%
           End If
         End If
       End If
       '
    Next KK%
    '
    If YAVERI% < 1 Then
      FIN& = LOF(NXQ%) / 4
      For KK% = 1 To LOF(NXQ%) / 4
        UU& = KK%
        Call TRACE(20, UU&, FIN&)
        Get #NXQ%, KK%, IVQ$
        CUE% = CVI(Mid$(IVQ$, 1, 2))
        If CUE% > 0 Then
          If YAPLA%(CUE%) < 1 Then
            NIVEI% = CVI(Mid$(IVQ$, 3, 2))
            If NIVEI% < 0 Then NIVEI% = 0
            If NIVEI% > 4 Then NIVEI% = 4
            '
            COCUE$ = CODCUE$(CUE%)
            If COCUE$ > UCUEPLAN$ Then
              ZZ$ = Space$(64)
              If NIVEI% >= 0 Then
                If NIVEI% <= 4 Then
                  NVX$ = Left$(String$(NIVEI%, ".") + TRIM$(Str$(NIVEI%)) + ".....", 5)
                  Call REPLA(ZZ$, NVX$, 1, 5)
                  Call REPLA(ZZ$, CODCUE$(CUE%), 10, 12)
                  Call REPLA(ZZ$, DENOCUE$(CUE%), 26, 32)
                  List1.AddItem ZZ$
                  YAPLA%(CUE%) = 1
                End If
              End If
            End If
          End If
        End If
      Next KK%
      YAVERI% = 1
    End If
    '
    List1.Refresh
    Screen.MousePointer = 1
    '
End Sub
'
Sub GUARDAPLAN()
   '
   Call BLOQARCH("CUENTAS")
   Call BLOQARCH("DENOCUE")
   Call BLOQARCH("INVECUE")
   Call BLOQARCH("PLAGRAL")
   Call BLOQARCH("PLANCUEN")
   Call BLOQARCH("CUECON")
   Call BLOQARCH("ECUECON")
   Call BLOQARCH("ECUECOT")
   '
   JJ& = 0: J4& = 0: J5& = 0: J6& = 0
   URECU& = ULTREG&("CUENTAS")
   Screen.MousePointer = 11
   REBLA$ = REGBLAN$("PLANCUEN")
   '
   OREC% = 0: QREC$ = TRIM$(Str$(Int(Timer)))
   FIN& = List1.ListCount
   List2.Clear
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     AGRECUE% = 0
     ZZ$ = List1.List(U& - 1)
     COCUE$ = TRIM$(Mid$(ZZ$, 9, 12))
       If COCUE$ = "" Then
         OREC% = OREC% + 1
         COCUE$ = "R-" + QREC$ + "-" + TRIM$(Str$(OREC%))
       End If
     DECUE$ = TRIM$(Mid$(ZZ$, 129, 128))
       If DECUE$ = "" Then
         DECUE$ = "Descripcion Recuperada " + COCUE$
       End If
     CUEIN% = Val(Mid$(ZZ$, 121, 8))
       If CUEIN% < 1 Or CUEIN% > URECU& Then
         URECU& = URECU& + 1
         CUEIN% = URECU&
         AGRECUE% = 1
       End If
     RECUENTA& = CUEIN%
     TICU$ = Mid$(ZZ$, 69, 1)
     IBLE% = 1: If UCase$(Mid$(ZZ$, 76, 1)) = "N" Then IBLE% = 0
     STAT% = 1: If UCase$(Mid$(ZZ$, 80, 1)) = "B" Then STAT% = 9
     NIVEI% = Val(TRIM$(REPLACAR$(TRIM$(Left$(ZZ$, 5)), ".", " ")))
     '
     ' GRABAR EN CUENTAS .DAT
     If AGRECUE% < 1 Then
         Z1$ = REGLEIDO$("CUENTAS", RECUENTA&)
       Else
         Z1$ = REGBLAN$("CUENTAS")
     End If
     Call REPLA(Z1$, COCUE$, 1, 12)
     Call REPLA(Z1$, TICU$, 13, 1)
     Call GRAREG("CUENTAS", Z1$, RECUENTA&)
     '
     ' GRABAR EN DENOCUE.DAT
     For KU& = 1 + ULTREG&("DENOCUE") To RECUENTA&
       Call GRAREG("DENOCUE", Space$(127) + ".", KU&)
     Next KU&
     Z2$ = REGLEIDO$("DENOCUE", RECUENTA&)
     Call REPLA(Z2$, DECUE$, 1, 100)
     Call REPLA(Z2$, DECUE$, 101, 28)
     Call GRAREG("DENOCUE", Z2$, RECUENTA&)
     '
     ' GRABAR EN PLANCUEN.DAT
     Z3$ = REBLA$
     Call REPLA(Z3$, COCUE$, 1, 12)
     Call REPLA(Z3$, DECUE$, 13, 46)
     Call REPLA(Z3$, MKI$(CUEIN%), 59, 2)
     Call REPLA(Z3$, Chr$(STAT%), 61, 1)
     Call REPLA(Z3$, Chr$(NIVEI%), 62, 1)
     Call REPLA(Z3$, TICU$, 63, 1)
     Call REPLA(Z3$, Chr$(IBLE%), 64, 1)
     Call REPLA(Z3$, DECUE$, 129, 128)
     JJ& = JJ& + 1
     Call GRAREG("PLANCUEN", Z3$, JJ&)
     '
     ' GRABAR EN PLAGRAL.DAT
     Z4$ = MKI$(CUEIN%) + MKI$(NIVEI%)
     Call GRAREG("PLAGRAL", Z4$, JJ&)
     '
     ' GRABAR EN LIST2 PARA INVECUE
     TTXX$ = Space$(18)
     Call REPLA(TTXX$, COCUE$, 1, 12)
     Call REPLA(TTXX$, TRIM$(Str$(CUEIN%)), 13, 6)
     For KI% = 1 To List2.ListCount
       If List2.List(KI% - 1) > TTXX$ Then
         List2.AddItem TTXX$, KI% - 1
         GoTo 25
       End If
     Next KI%
     List2.AddItem TTXX$
     '
25   Z4$ = REGBLAN$("CUECON")
     Z5$ = REGBLAN$("ECUECON")
     '
     Call REPLA(Z5$, COCUE$, 1, 12)
     Call REPLA(Z5$, DECUE$, 13, 28)
     Call REPLA(Z5$, MKI$(CUEIN%), 41, 2)
     J5& = J5& + 1
     Call GRAREG("ECUECOT", Z5$, J5&)
     '
     If STAT = 1 Then
       If IBLE% = 1 Then
         '
         J6& = J6& + 1
         Call GRAREG("ECUECON", Z5$, J6&)
         '
         If TICU$ = "I" Or TICU$ = "G" Then
           Call REPLA(Z4$, MKI$(CUEIN%), 1, 2)
           Call REPLA(Z4$, DECUE$, 3, 30)
           Call REPLA(Z4$, COCUE$, 33, 12)
           Call REPLA(Z4$, COCUE$, 45, 12)
           J4& = J4& + 1
           Call GRAREG("CUECON", Z4$, J4&)
         End If
         '
       End If
     End If
     '
     
   Next U&
   '
   Call SETULTREG("PLANCUEN", JJ&)
   Call SETULTREG("PLAGRAL", JJ&)
   Call SETULTREG("CUECON", J4&)
   Call SETULTREG("ECUECOT", J5&)
   Call SETULTREG("ECUECON", J6&)
   Call CIERRARCH("CUENTAS")
   Call CIERRARCH("DENOCUE")
   '
   REBLA$ = REGBLAN$("INVECUE")
   J7& = 0
   For KI% = 1 To List2.ListCount
     TTXX$ = List2.List(KI% - 1)
     COCUE$ = Left$(TTXX$, 12)
     ICUE% = Val(Mid$(TTXX$, 13))
     Z7$ = REBLA$
     Call REPLA(Z7$, COCUE$, 1, 12)
     Call REPLA(Z7$, MKI$(ICUE%), 13, 2)
     J7& = J7& + 1
     Call GRAREG("INVECUE", Z7$, J7&)
   Next KI%
   Call SETULTREG("INVECUE", J7&)
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Close
   '
   ABCT$ = LOADFILE$(LUDAT$ + "CUENTAS.DAT"): Call SAVEFILE(LUDAT$ + "CUENTAS.OLD", ABCT$)
   ABCT$ = LOADFILE$(LUDAT$ + "DENOCUE.DAT"): Call SAVEFILE(LUDAT$ + "DENOCUE.OLD", ABCT$)
   ABCT$ = LOADFILE$(LUDAT$ + "INVECUE.DAT"): Call SAVEFILE(LUDAT$ + "INVECUE.OLD", ABCT$)
   ABCT$ = LOADFILE$(LUDAT$ + "PLAGRAL.DAT"): Call SAVEFILE(LUDAT$ + "PLAGRAL.OLD", ABCT$)
   '
   Screen.MousePointer = 1
   '
End Sub

Sub CARGAPLAN()
   '
   Screen.MousePointer = 11
   List1.Clear
   If ULTREG("PLANCUEN") < 1 Then
     ' CARGA DE PLAGRAL
     FIN& = ULTREG&("PLAGRAL")
     For U& = 1 To FIN&
       IVP$ = REGLEIDO$("PLAGRAL", U&)
       '
       CUE% = CVI(Mid$(IVP$, 1, 2))
       NIVEI% = CVI(Mid$(IVP$, 3, 2))
       If NIVEI% < 0 Then NIVEI% = 0
       If NIVEI% > 4 Then NIVEI% = 4
       '
       If CUE% > 0 Then
         ZZ$ = Space$(256)
         If NIVEI% >= 0 Then
           If NIVEI% <= 4 Then
             NVX$ = Left$(String$(NIVEI%, ".") + TRIM$(Str$(NIVEI%)) + ".....", 5)
             COCUE$ = CODCUE$(CUE%)
             DECUE$ = DENOCUE$(CUE%)
             '
             Call REPLA(ZZ$, NVX$, 1, 5)
             Call REPLA(ZZ$, COCUE$, 9, 12)
             Call REPLA(ZZ$, DECUE$, 23, 44)
             Call REPLA(ZZ$, DECUE$, 129, 128)
             '
             RECUEN& = CUE%
             XX$ = REGLEIDO$("CUENTAS", RECUEN&)
             TICUEX$ = UCase$(Mid$(XX$, 13, 1))
             If TICUEX$ = "H" Then
                 TICUEX$ = "Haber"
               ElseIf TICUEX$ = "I" Then
                 TICUEX$ = "Ingresos"
               ElseIf TICUEX$ = "G" Then
                 TICUEX$ = "Gastos"
               Else
                 TICUEX$ = "Debe"
             End If
             Call REPLA(ZZ$, TICUEX$, 69, 5)
             '
             IMPUBLE$ = ""
             Call REPLA(ZZ$, IMPUBLE$, 76, 2)
             '
             STATUX$ = ""
             Call REPLA(ZZ$, STATUX$, 80, 4)
             '
             NUINT$ = AJUSTD$(Str$(CUE%), 8)
             Call REPLA(ZZ$, NUINT$, 121, 8)
             '
             List1.AddItem ZZ$
             '
           End If
         End If
       End If
     Next U&
     GoTo 90
   End If
   '
   '
   ' CARGA DE PLANCUEN
   FIN& = ULTREG&("PLANCUEN")
   For U& = 1 To FIN&
     XX$ = REGLEIDO$("PLANCUEN", U&)
     ZZ$ = Space$(256)
     '
     COCUE$ = Left$(XX$, 12)
     DECUE$ = Mid$(XX$, 129, 128)
     NIVEI% = Asc(Mid$(XX$, 62, 1))
     NVX$ = Left$(String$(NIVEI%, ".") + TRIM$(Str$(NIVEI%)) + ".....", 5)
     '
     Call REPLA(ZZ$, NVX$, 1, 5)
     Call REPLA(ZZ$, COCUE$, 9, 12)
     Call REPLA(ZZ$, DECUE$, 23, 44)
     Call REPLA(ZZ$, DECUE$, 129, 128)
     '
     TICUEX$ = UCase$(Mid$(XX$, 63, 1))
     If TICUEX$ = "H" Then
         TICUEX$ = "Haber"
       ElseIf TICUEX$ = "I" Then
         TICUEX$ = "Ingresos"
       ElseIf TICUEX$ = "G" Then
         TICUEX$ = "Gastos"
       Else
         TICUEX$ = "Debe"
     End If
     Call REPLA(ZZ$, TICUEX$, 69, 5)
     '
     IMPUBLE$ = ""
     If Asc(Mid$(XX$, 64, 1)) < 1 Then IMPUBLE$ = "No"
     Call REPLA(ZZ$, IMPUBLE$, 76, 2)
     '
     STATUX$ = ""
     If Asc(Mid$(XX$, 61, 1)) = 9 Then STATUX$ = "Baja"
     Call REPLA(ZZ$, STATUX$, 80, 4)
     '
     NUINT$ = AJUSTD$(Str$(CVI(Mid$(XX$, 59, 2))), 8)
     Call REPLA(ZZ$, NUINT$, 121, 8)
     '
     List1.AddItem ZZ$
     '
   Next U&
   '
90 List1.Refresh
   Screen.MousePointer = 1
   '
End Sub
'
Sub RECUECON()
   '
   Screen.MousePointer = 11
   Dim FCUEN As String * 64
   Dim YAPLAN%(32767, 2)
   Dim IVP As String * 4
   NOHAYPLAN% = 0
   '
   EJERCI$ = EJACT$
   LU$ = LUDAT$
   ACUEN% = FILEOPEN%(LU$ + "CUENTAS." + EXTE$, 0, 64)
   For UI% = LOF(ACUEN%) / 64 To 1 Step -1
       Get #ACUEN%, UI%, FCUEN$
       If Asc(Left$(FCUEN$, 1)) > 0 Then
          NUCUEN% = UI%
          GoTo 5
       End If
   Next UI%
   NUCUEN% = 0
   '
5  Close #ACUEN%: ACUEN% = 0
   EXTEPLAN$ = "DAT": IPLAN% = 1: LPLAN$ = "Actual"
   '
6  NXP% = FILEOPEN%(LUDAT$ + "PLAGRAL." + EXTEPLAN$, 0, 4)
   If LOF(NXP%) < 4 Then NOHAYPLAN% = NOHAYPLAN% = NOHAYPLAN% + 1
   '
   FIN& = LOF(NXP%) / 4
   For U& = 1 To LOF(NXP%) / 4
       Call TRACE(20, U&, FIN&)
       Get #NXP%, U&, IVP$
       ICUE% = CVI(Left$(IVP$, 2))
       If ICUE% > 0 Then
          If ICUE% <= 32767 Then
             YAPLAN%(ICUE%, IPLAN%) = YAPLAN%(ICUE%, IPLAN%) + 1
          End If
       End If
   Next U&
   Close #NXP%: NXP% = 0
   '
   For KKI% = 1 To NUCUEN%
       If YAPLAN%(KKI%, IPLAN%) > 1 Then
          Call MENSERR(24, "Cuenta '" + TRIM$(CODCUE$(KKI%)) + "'\Duplicada en el Plan de Cuentas Ejercicio " + LPLAN$ + ".")
       End If
   Next KKI%
   '
   If EXTEPLAN$ = "DAT" Then
       EXTEPLAN$ = "OLD"
       IPLAN% = 2
       LPLAN$ = "Anterior"
       GoTo 6
   End If
   '
   K2& = 0
   FIN& = ULTREG&("ECUECOT")
   For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
      Z2$ = REGLEIDO$("ECUECOT", U&)
      ICUE% = CVI(Mid$(Z2$, 41, 2))
      If ICUE% > 0 Then
         If ICUE% <= NUCUEN% Then
            If YAPLAN%(ICUE%, 1) Or YAPLAN%(ICUE%, 2) > 0 Or NOHAYPLAN% = 2 Then
               K2& = K2& + 1
               Call GRAREG("ECUECON", Z2$, K2&)
            End If
         End If
      End If
   Next U&
   '
   Call SETULTREG("ECUECON", K2&)
   Call CIERRARCH("ECUECON")
   Call CIERRARCH("ECUECOT")
   Call FRESHECHO("ECUECON")
   Screen.MousePointer = 1
   '
End Sub

Private Sub Text1_DblClick()
   If TRIM$(Text1) <> "" Then
     MAYUS% = 0
     If UCase(TRIM$(Text1)) = TRIM$(Text1) Then MAYUS% = 1
     '
     If MAYUS% = 0 Then
         Text1 = UCase$(Text1)
       Else
         TTXX$ = " " + LCase$(Text1)
         For KKI% = 2 To Len(TTXX$)
           If Mid$(TTXX$, KKI% - 1, 1) = " " Then
             Mid$(TTXX$, KKI%, 1) = UCase$(Mid$(TTXX$, KKI%, 1))
           End If
         Next KKI%
         Text1 = Mid$(TTXX$, 2)
     End If
   End If
End Sub

Private Sub Text1_GotFocus()
   If TRIM$(Text11) = "" Then
      Command21.SetFocus
   End If
End Sub

Private Sub Text11_GotFocus()
   If TRIM$(Text11) = "" Then
     On Error Resume Next
     Command21.SetFocus
     On Error GoTo 0
   End If
End Sub

Private Sub Text14_GotFocus()
   If TRIM$(Text11) = "" Or Val(Label8) < 1 Then
      Command21.SetFocus
   End If
End Sub

Private Sub VScroll1_Change()
   Label2 = 4 - VScroll1.Value
End Sub

Sub NUEVACUECON()
   VDEF$ = Space$(64)
10 OBX$ = OBJPLA$("NUEVACUECON", VDEF$)
   If OBX$ <> "" Then
     VDEF$ = OBX$
     CONUCUE$ = TRIM$(Left$(OBX$, 12))
     DENUCUE$ = TRIM$(Mid$(OBX$, 13, 48))
     If Len(CONUCUE$) < 1 Then
       Call MENSERR(24, "Falta Codigo de Cuenta")
       GoTo 10
     End If
     '
     FIN& = List1.ListCount
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       If UCase$(TRIM$(Mid$(List1.List(U& - 1), 9, 12))) = UCase$(CONUCUE$) Then
         Call MENSERR(24, "Cuenta Registrada en el Plan de Cuentas")
         List1.ListIndex = U& - 1
         Exit Sub
       End If
     Next U&
     '
     Text11 = CONUCUE$
     Text1 = DENUCUE$
     Option1(0).Value = True
     Option2(0).Value = True
     Option3(0).Value = True
     Label8 = "0"
     If List1.ListIndex <= 0 Then
         NVLX% = 0
       Else
         NVLX% = Val(TRIM$(REPLACAR$(TRIM$(Left$(List1.TEXT, 5)), ".", " ")))
     End If
     If NVLX% > 4 Then NVLX% = 4
     Label2 = TRIM$(Str$(NVLX%))
     VScroll1.Value = 4 - NVLX%
     '
     On Error Resume Next
     Text11.SetFocus
     On Error GoTo 0
     '
   End If
     
End Sub
   
Sub AGRECUENTA(MODOI%)
   ' MODOI%=1 --> Append
   ' MODOI%=2 --> Insert
   '
   COCUE$ = TRIM$(Text11)
   If COCUE$ = "" Then Exit Sub
   '
   FIN& = List1.ListCount
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     If UCase$(TRIM$(Mid$(List1.List(U& - 1), 9, 12))) = UCase$(COCUE$) Then
       Call MENSERR(24, "Cuenta Existente en el Plan de Cuentas")
       List1.ListIndex = U& - 1
       Exit Sub
     End If
   Next U&
   '
   NIVEI% = Val(Label2)
   If NIVEI% < 0 Or NIVEI% > 4 Then
      Call MENSERR(24, "Mal Nivel de Agrupación")
      Exit Sub
   End If
   '
   ' aqui validar el nivel de agrupacion segun cuenta anterior
   '
   ZZ$ = Space$(256)
   NVX$ = Left$(String$(NIVEI%, ".") + TRIM$(Str$(NIVEI%)) + ".....", 5)
   Call REPLA(ZZ$, NVX$, 1, 5)
   Call REPLA(ZZ$, Text11, 9, 12)
   Call REPLA(ZZ$, Text1, 23, 44)
   Call REPLA(ZZ$, Text1, 129, 128)
   '
   TICUE$ = "Debe"
   For KKI% = 0 To 3
     If Option1(KKI%).Value = True Then TICUE$ = Option1(KKI%).Caption
   Next KKI%
   Call REPLA(ZZ$, TICUE$, 69, 5)
   '
   IMPUBLE$ = ""
   If Option2(1).Value = True Then IMPUBLE$ = "No"
   Call REPLA(ZZ$, IMPUBLE$, 76, 2)
   '
   STATUX$ = ""
   If Option3(1).Value = True Then STATUX$ = "Baja"
   Call REPLA(ZZ$, STATUX$, 80, 4)
   '
   NUINT$ = AJUSTD$(Label8, 8)
   Call REPLA(ZZ$, NUINT$, 121, 8)
   '
   If List1.ListIndex < 0 Then MODOI% = 1   ' SI NO DA ERROR
   '
   If MODOI% = 1 Then
       List1.AddItem ZZ$
       On Error Resume Next
       List1.ListIndex = List1.ListCount - 1
       On Error GoTo 0
     Else
       ILISTA& = List1.ListIndex
       List1.AddItem ZZ$, ILISTA&
       On Error Resume Next
       List1.ListIndex = ILISTA&
       On Error GoTo 0
   End If
   '
   MODOEDIT% = 0
   ILISTACT& = 0
   ACTIVLIST$ = ""
   Command4.Enabled = True
   Command8.Enabled = True
   Command24.Enabled = True
   '
   Text1.TEXT = ""
   Text11.TEXT = ""
   Text14 = ""
   '
   Label2 = ""
   Label8 = ""
   Option1(0) = True
   Option2(0) = True
   Option3(0) = True
   '
End Sub

Sub EDITACUEN()
   '
   If MODOEDIT% > 0 Then
     Call MENSERR(24, "Primero Guarde la Cuenta Editada")
     Exit Sub
   End If
   '
   U& = List1.ListIndex
   If U& < 0 Then Exit Sub
   If U& > List1.ListCount - 1 Then Exit Sub
   '
   MODOEDIT% = 1
   ILISTACT& = U&
   ACTIVLIST$ = List1.List(U&)
   Command4.Enabled = False
   Command8.Enabled = False
   Command24.Enabled = False
   '
   URECU& = ULTREG&("CUENTAS")
   ZZ$ = List1.List(U&)
   Text11 = TRIM$(Mid$(ZZ$, 9, 12))
   Text1 = TRIM$(Mid$(ZZ$, 129, 128))
     CUEIN% = Val(Mid$(ZZ$, 121, 8))
     If CUEIN% < 1 Or CUEIN% > URECU& Then CUEIN% = 0
   Label8 = TRIM$(Str$(CUEIN%))
     TICU$ = Mid$(ZZ$, 69, 1)
   For KKI% = 0 To 3
     If Left$(Option1(KKI%).Caption, 1) = TICU$ Then Option1(KKI%).Value = True
   Next KKI%
   '
   Option2(0).Value = True: If UCase$(Mid$(ZZ$, 76, 1)) = "N" Then Option2(1).Value = True
   Option3(0).Value = True: If UCase$(Mid$(ZZ$, 80, 1)) = "B" Then Option3(1).Value = True
   '
   NIVEI% = Val(TRIM$(REPLACAR$(TRIM$(Left$(ZZ$, 5)), ".", " ")))
   Label2 = TRIM$(Str$(NIVEI%))
   VScroll1.Value = 4 - NIVEI%
   '
   List1.RemoveItem U&
   If U& < List1.ListCount Then
      List1.ListIndex = U&
   End If
   '
End Sub
