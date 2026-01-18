VERSION 5.00
Begin VB.Form FORNOCON 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes de No-Conformidad"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Material"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Index           =   1
      Left            =   105
      TabIndex        =   22
      Top             =   105
      Width           =   7890
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1080
         Left            =   105
         TabIndex        =   23
         Top             =   300
         Width           =   7740
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo:"
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
            Left            =   480
            TabIndex        =   31
            Top             =   105
            Width           =   1005
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Descripcion:"
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
            Left            =   105
            TabIndex        =   30
            Top             =   525
            Width           =   1380
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6825
            TabIndex        =   29
            Top             =   0
            Width           =   750
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "Interno:"
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
            Left            =   5670
            TabIndex        =   28
            Top             =   105
            Width           =   1005
         End
         Begin VB.Label Label39 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4830
            TabIndex        =   27
            Top             =   0
            Width           =   540
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "U-M:"
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
            Left            =   3675
            TabIndex        =   26
            Top             =   105
            Width           =   1005
         End
         Begin VB.Label Label6 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1575
            TabIndex        =   25
            Top             =   525
            Width           =   6000
         End
         Begin VB.Label Label37 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1575
            TabIndex        =   24
            Top             =   0
            Width           =   2115
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
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
      Height          =   2025
      Left            =   10815
      TabIndex        =   3
      Top             =   2835
      Width           =   990
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
         Height          =   750
         Left            =   105
         Picture         =   "FORNOCON.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Nuevo Informe de No-Conformidad"
         Top             =   315
         Width           =   750
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Restore"
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
         Left            =   105
         Picture         =   "FORNOCON.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Restablece los Valores Anteriores"
         Top             =   1050
         Width           =   750
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Varios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   10815
      TabIndex        =   5
      Top             =   4935
      Width           =   990
      Begin VB.CommandButton Command37 
         Caption         =   "Setup"
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
         Left            =   105
         Picture         =   "FORNOCON.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   300
         Width           =   750
      End
      Begin VB.CommandButton Command39 
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
         Height          =   750
         Left            =   105
         Picture         =   "FORNOCON.frx":0896
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Imprime Plan de Inspección"
         Top             =   1050
         Width           =   750
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Seguimiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6600
      Index           =   0
      Left            =   105
      TabIndex        =   9
      Top             =   1680
      Width           =   10620
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   6
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   44
         Top             =   5880
         Width           =   8850
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   5
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         Top             =   5250
         Width           =   8850
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   4
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   4620
         Width           =   8850
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   3
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   3465
         Width           =   8850
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   2
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Top             =   2415
         Width           =   8850
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   1
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   1365
         Width           =   8850
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   0
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   315
         Width           =   8850
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         X1              =   0
         X2              =   10605
         Y1              =   4530
         Y2              =   4530
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000F&
         X1              =   0
         X2              =   10605
         Y1              =   4515
         Y2              =   4515
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   6
         Left            =   105
         TabIndex        =   45
         Top             =   5880
         Width           =   1485
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cierre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   5
         Left            =   105
         TabIndex        =   43
         Top             =   5250
         Width           =   1485
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Seguimiento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   4
         Left            =   105
         TabIndex        =   41
         Top             =   4620
         Width           =   1485
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Acción Correctiva:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   3
         Left            =   105
         TabIndex        =   39
         Top             =   3465
         Width           =   1485
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Resolución de No-Conformidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   2
         Left            =   105
         TabIndex        =   37
         Top             =   2415
         Width           =   1485
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "  Disposición o Propuesta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   1
         Left            =   105
         TabIndex        =   35
         Top             =   1365
         Width           =   1485
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "    Descripción de la No-Conformidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   0
         Left            =   105
         TabIndex        =   33
         Top             =   315
         Width           =   1485
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Informe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Index           =   0
      Left            =   8085
      TabIndex        =   10
      Top             =   105
      Width           =   2640
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1080
         Left            =   105
         TabIndex        =   11
         Top             =   360
         Width           =   2490
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
            Height          =   280
            Index           =   0
            Left            =   2100
            Picture         =   "FORNOCON.frx":0F00
            TabIndex        =   12
            Top             =   0
            Width           =   175
         End
         Begin VB.Label Label9 
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
            Height          =   300
            Left            =   840
            TabIndex        =   21
            Top             =   740
            Width           =   1470
         End
         Begin VB.Label Label8 
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
            Height          =   300
            Left            =   1260
            TabIndex        =   20
            Top             =   370
            Width           =   1050
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   1260
            TabIndex        =   19
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "Origen:"
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
            Left            =   -210
            TabIndex        =   18
            Top             =   845
            Width           =   1005
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "Informe No:"
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
            Left            =   105
            TabIndex        =   14
            Top             =   105
            Width           =   1005
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
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
            Left            =   105
            TabIndex        =   13
            Top             =   475
            Width           =   1005
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Main"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   10815
      TabIndex        =   0
      Top             =   105
      Width           =   990
      Begin VB.CommandButton Command3 
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
         Left            =   105
         Picture         =   "FORNOCON.frx":120A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Abandona Ficha Plan de Inspección"
         Top             =   315
         Width           =   750
      End
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
         Height          =   750
         Left            =   105
         Picture         =   "FORNOCON.frx":1354
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   1050
         Width           =   750
      End
      Begin VB.CommandButton Command4 
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
         Left            =   105
         Picture         =   "FORNOCON.frx":165E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Ayuda en Línea"
         Top             =   1785
         Width           =   750
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10350
      Picture         =   "FORNOCON.frx":1968
      Top             =   7750
      Width           =   2010
   End
   Begin VB.Label DOCUORI 
      Height          =   225
      Left            =   10815
      TabIndex        =   16
      Top             =   7035
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label COMMLAB 
      Height          =   225
      Left            =   10815
      TabIndex        =   15
      Top             =   6825
      Visible         =   0   'False
      Width           =   990
   End
End
Attribute VB_Name = "FORNOCON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODIPLA%
Dim GRABATODO%
'
Private Sub Command1_Click()
   Command1.Enabled = False
   GRABATODO% = 1
   Call GRAINFONOCO
   GRABATODO% = 0
   Command1.Enabled = True
End Sub


Private Sub Command2_Click()
   Call GRAINFONOCO
   Call CARINFONOCO
End Sub

Private Sub COMMAND21_CLICK()
   '
   Call GRAINFONOCO
   HOII% = HOY(HO$)
   VDEF$ = String$(4, 0) + Space$(12) + MKI$(0)
   UREINF& = ULTREG&("INFONOCO")
   If UREINF& >= 32767 Then
     Call MENSERR(24, "Capacidad de la Base de Datos Saturada")
     Call FINAL("")
   End If
   NINF% = 1 + UREINF&
   Call REPLA(VDEF$, MKI$(NINF%), 1, 2)
   Call REPLA(VDEF$, MKI$(HOII%), 3, 2)
   '
10 Screen.MousePointer = 1
   OBX$ = OBJPLA$("NUEINFONOCO", VDEF$)
   If OBX$ = "" Then
     Exit Sub
   End If
   '
   VDEF$ = OBX$
   ORINF$ = TRIM$(Mid$(OBX$, 5, 12))
   If ORINF$ = "" Then
      Call MENSERR(24, "Falta Origen del Informe")
      GoTo 10
   End If
   '
   CII% = CVI(Mid$(OBX$, 17, 2))
   If TRIM$(CODEXT$(CII%)) = "" Then
      Call MENSERR(24, "Código no Válido")
      GoTo 10
   End If
   '
   FEX = CVI(Mid$(OBX$, 3, 2))
   If FEX < 1 Then
      Call MENSERR(24, "Fecha de Generación no Válida")
      GoTo 10
   End If
   '
   Call BLOQARCH("INFONOCO")
   TTXX$ = Space$(80)
   Call REPLA(TTXX$, MKI$(NINF%), 1, 2)
   Call REPLA(TTXX$, FECHATEX$(FEX), 3, 8)
   Call REPLA(TTXX$, ORINF$, 13, 12)
   Call REPLA(TTXX$, CODEXT$(CII%), 27, 16)
   Call REPLA(TTXX$, DESCRIT0$(CII%), 43, 38)
   Call REPLA(TTXX$, MKI$(CII%), 81, 2)
   Call REGAPP("INFONOCO", TTXX$)
   Call CIERRARCH("INFONOCO")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   Label5 = TRIM$(Str$(NINF%))
   '
End Sub

Private Sub Command3_Click()
    Call GRAINFONOCO
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command37_Click()
   Command37.Enabled = False
   FORDOCINGE.Show 1
   Command37.Enabled = True
End Sub

Private Sub Command39_Click()
   Command39.Enabled = False
   DOCUORI.Caption = ""
   COMMLAB.Caption = "PRINT"
   Command39.Enabled = True
End Sub

Private Sub Command4_Click()
   Call MENSERR(24, "Ayuda no Disponible")
End Sub

Private Sub Command9_Click(Index As Integer)
   Call GRAINFONOCO
   Call SELECHO("INFONOCO")
   VDEV$ = VALACT1$("INFONOCO")
   If VDEV$ <> "" Then
      Label5 = VDEV$
   End If
End Sub

Private Sub COMMLAB_Change()
   If TRIM$(UCase$(COMMLAB.Caption)) = "PRINT" Then
      DOCUOR$ = TRIM$(DOCUORI.Caption)
      Call PRINOCONF(DOCUOR$)
      COMMLAB.Caption = ""
      DOCUORI.Caption = ""
   End If
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + " - " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
End Sub

Private Sub Form_KeyDown(Keycode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Resize()
   DoEvents
   If WindowState = 1 Then
      Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Sub GRAINFONOCO()
   '
   NINF% = Val(Label5)
   If NINF% < 1 Or NINF% > ULTREG&("INFONOCO") Then
     Exit Sub
   End If
   UREDETA& = ULTREG&("DETANOCO")
   '
   If MODIPLA% <> 1 Then Exit Sub
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se ha Modificado el Informe de No Conformidad Nro." + TRIM$(Label5) + "\\Registrar Cambios\Cancelar") <> 1 Then
       Call CARINFONOCO
       MODIPLA% = 0
       Exit Sub
     End If
   End If
   '
   Screen.MousePointer = 11
   Call BLOQARCH("DETANOCO")
   RNC$ = RECFILT$("DETANOCO", 1, MKI$(NINF%))
   KK& = (-3)
   '
   For KKI% = 0 To 6
      TTXX$ = RTrim$(Text2(KKI%))
10    TTY1$ = Left$(TTXX$, 124)
      XX$ = MKI$(NINF%) + MKI$(KKI%) + Space$(124)
      Call REPLA(XX$, TTY1$, 5, 124)
      '
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           UREDETA& = UREDETA& + 1
           RENOTA& = UREDETA&
      End If
      Call GRAREG("DETANOCO", XX$, RENOTA&)
      TTXX$ = Mid$(TTXX$, 125)
      If Len(TTXX$) > 0 Then GoTo 10
      '
   Next KKI%
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("DETANOCO", MKI$(0) + MKI$(0) + Space$(124), RENOTA&)
   Wend
   '
   Call CIERRARCH("DETANOCO")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   MODIPLA% = 0
   Screen.MousePointer = 1
   '
End Sub


Sub PRINOCONF(DOCUORIGEN$)
  '
  FORIPRE$ = TRIM$(CONTROL$("FORGRAFI", "DOCUGRA5"))

  If FORIPRE$ = "" Then
       PRINTFORM
     Else
       Screen.MousePointer = 11
       FEX = HOY(HOS$)
       FECHDOC$ = HOS$
       DOCUORIG$ = Label9
       NUMEDOC$ = TRIM$(Label5)
       TIPODOC$ = "Informe de No-Conformidad"
       '
       CODPARAM$(1) = "FECH-LAN"
       CODPARAM$(2) = "COD-MAT"
       CODPARAM$(3) = "DES-MAT"
       CODPARAM$(4) = "REF-MAT1"
       CODPARAM$(5) = "REF-MAT2"
       CODPARAM$(6) = "REF-MAT3"
       CODPARAM$(7) = "REF-CONJ"
       CODPARAM$(8) = "COL-CONJ"
       CODPARAM$(9) = "TAL-CONJ"
       CODPARAM$(10) = "MEDIDA-1"
       CAPARDOC% = 10
       '
       DOCPARAM$(1) = TRIM$(Label8)
       DOCPARAM$(2) = TRIM$(Label37)
       DOCPARAM$(3) = TRIM$(Label6)
       DOCPARAM$(4) = RTrim$(Text2(0))
       DOCPARAM$(5) = RTrim$(Text2(1))
       DOCPARAM$(6) = RTrim$(Text2(2))
       DOCPARAM$(7) = RTrim$(Text2(3))
       DOCPARAM$(8) = RTrim$(Text2(4))
       DOCPARAM$(9) = RTrim$(Text2(5))
       DOCPARAM$(10) = RTrim$(Text2(6))
       '
       CACOLTAB1% = 0
       '
       Call PRINTDOC("@" + FORIPRE$)
       Screen.MousePointer = 1
       '
  End If
  '
End Sub

Private Sub Label5_Change()
   Call CARINFONOCO
End Sub

Sub CARINFONOCO()
   '
   NINF% = Val(Label5)
   URN& = ULTREG&("INFONOCO")
   If NINF% < 1 Or NINF% > URN& Then
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   REINFO& = NINF%
   YY$ = REGLEIDO$("INFONOCO", REINFO&)
   Label8 = Mid$(YY$, 3, 8)
   Label9 = Mid$(YY$, 12, 12)
   CII% = CVI(Mid$(YY$, 81, 2))
   Label28 = TRIM$(Str$(CII%))
   Label37 = CODEXT$(CII%)
   Label39 = UNIMED$(CII%)
   Label6 = DESCRIT0$(CII%)
   '
   For KKI% = 0 To 6
      Text2(KKI%) = ""
   Next KKI%
   '
   RNC$ = RECFILT$("DETANOCO", 1, MKI$(NINF%))
   UREDETA& = ULTREG&("DETANOCO")
   '
   For U& = 1 To Len(RNC$) Step 4
     RENOTA& = CVS(Mid$(RNC$, U&, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= UREDETA& Then
         Y$ = REGLEIDO$("DETANOCO", RENOTA&)
         If CVI(Left$(Y$, 2)) = NINF% Then
           CAPI% = CVI(Mid$(Y$, 3, 2))
           If CAPI% >= 0 Then
             If CAPI% <= 6 Then
               Text2(CAPI%) = Text2(CAPI%) + Mid$(Y$, 5)
             End If
           End If
         End If
       End If
     End If
   Next U&
   '
   For KKI% = 0 To 6
      TTXX$ = RTrim(Text2(KKI%))
      Text2(KKI%) = TTXX$
   Next KKI%
   '
   MODIPLA% = 0
   OPENFORMS = DoEvents
   Screen.MousePointer = 1
   '
End Sub
'

Private Sub Text2_Change(Index As Integer)
   MODIPLA% = 1
End Sub

Private Sub Text2_GotFocus(Index As Integer)
   If Val(Label5) < 1 Then
      Command9(0).SetFocus
      Exit Sub
   End If
End Sub
