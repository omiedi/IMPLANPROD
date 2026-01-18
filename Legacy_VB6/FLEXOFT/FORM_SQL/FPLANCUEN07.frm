VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FPLANCUEN07 
   BackColor       =   &H0095C5E1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plan General de Cuentas"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   240
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   36
      Top             =   960
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   10800
      TabIndex        =   35
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0095C5E1&
      Caption         =   "EDICIÓN CUENTA ACTIVA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   120
      TabIndex        =   7
      Top             =   5880
      Width           =   11550
      Begin VB.Frame Frame2 
         BackColor       =   &H0095C5E1&
         Caption         =   "STATUS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   850
         Left            =   6720
         TabIndex        =   24
         Top             =   1200
         Width           =   975
         Begin VB.OptionButton Option3 
            Appearance      =   0  'Flat
            BackColor       =   &H0095C5E1&
            Caption         =   "Baja"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   80
            TabIndex        =   26
            Top             =   530
            Width           =   735
         End
         Begin VB.OptionButton Option3 
            Appearance      =   0  'Flat
            BackColor       =   &H0095C5E1&
            Caption         =   "Activa"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   80
            TabIndex        =   25
            Top             =   290
            Width           =   850
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         TabIndex        =   22
         Top             =   480
         Width           =   4005
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
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
         TabIndex        =   20
         Text            =   "FPLANCUEN07.frx":0000
         Top             =   1290
         Width           =   6045
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H0095C5E1&
         Caption         =   "IMPUTABLE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   8640
         TabIndex        =   17
         Top             =   390
         Width           =   1695
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H0095C5E1&
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   19
            Top             =   290
            Width           =   615
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H0095C5E1&
            Caption         =   "Si"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   18
            Top             =   290
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H0095C5E1&
         Caption         =   "TIPO DE CUENTA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   850
         Left            =   7800
         TabIndex        =   12
         Top             =   1200
         Width           =   2535
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H0095C5E1&
            Caption         =   "Ingresos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   15
            Top             =   290
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H0095C5E1&
            Caption         =   "Gastos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   1320
            TabIndex        =   16
            Top             =   530
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H0095C5E1&
            Caption         =   "Haber"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   14
            Top             =   530
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H0095C5E1&
            Caption         =   "Debe"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   13
            Top             =   290
            Width           =   1215
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   560
         Left            =   8160
         Max             =   5
         TabIndex        =   10
         Top             =   480
         Width           =   220
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Editar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10580
         Picture         =   "FPLANCUEN07.frx":0002
         TabIndex        =   4
         ToolTipText     =   "Edición de la Cuenta Activa"
         Top             =   1650
         Width           =   825
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10580
         Picture         =   "FPLANCUEN07.frx":030C
         TabIndex        =   3
         ToolTipText     =   "Borra la Cuenta Activa del Plan"
         Top             =   1260
         Width           =   825
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Insertar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10580
         Picture         =   "FPLANCUEN07.frx":0616
         TabIndex        =   2
         ToolTipText     =   "Intercala Cuenta en el Plan "
         Top             =   870
         Width           =   825
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
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
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10580
         Picture         =   "FPLANCUEN07.frx":0920
         TabIndex        =   1
         ToolTipText     =   "Agrega Cuenta al Final del Plan"
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7000
         TabIndex        =   29
         Top             =   780
         Width           =   555
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C-I"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6840
         TabIndex        =   28
         Top             =   550
         Width           =   360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   10440
         X2              =   10440
         Y1              =   120
         Y2              =   2200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2520
         TabIndex        =   27
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7440
         TabIndex        =   23
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   21
         Top             =   1050
         Width           =   1920
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   560
         Left            =   7800
         TabIndex        =   11
         Top             =   480
         Width           =   405
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   8
         Top             =   480
         Width           =   960
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0095C5E1&
      BorderStyle     =   0  'None
      Height          =   5480
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   10360
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   5460
         Index           =   1
         Left            =   2560
         TabIndex        =   34
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   5460
         Index           =   4
         Left            =   9000
         TabIndex        =   33
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   5460
         Index           =   3
         Left            =   8160
         TabIndex        =   32
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   5460
         Index           =   2
         Left            =   9480
         TabIndex        =   31
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   5460
         Index           =   0
         Left            =   960
         TabIndex        =   30
         Top             =   0
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
         ItemData        =   "FPLANCUEN07.frx":0C2A
         Left            =   0
         List            =   "FPLANCUEN07.frx":0C2C
         TabIndex        =   6
         Top             =   480
         Width           =   10360
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   500
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   10335
         TabIndex        =   9
         Top             =   0
         Width           =   10360
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   $"FPLANCUEN07.frx":0C2E
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   0
            TabIndex        =   38
            Top             =   120
            Width           =   10305
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FPLANCUEN07.frx":0CF0
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   360
      OleObjectBlob   =   "FPLANCUEN07.frx":0F24
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H002C3D46&
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   10750
      ScaleHeight     =   5745
      ScaleWidth      =   1185
      TabIndex        =   39
      Top             =   0
      Width           =   1215
      Begin VB.CommandButton Command21 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   120
         Picture         =   "FPLANCUEN07.frx":0F9E
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Alta de Nueva Cuenta Contable"
         Top             =   2655
         Width           =   750
      End
      Begin VB.CommandButton Command24 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   120
         Picture         =   "FPLANCUEN07.frx":10E8
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Anulación de Cuenta Contable"
         Top             =   3165
         Width           =   750
      End
      Begin VB.CommandButton Command6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   120
         Picture         =   "FPLANCUEN07.frx":13F2
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Imprime Listado de Cuentas Contables "
         Top             =   4155
         Width           =   750
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   120
         Picture         =   "FPLANCUEN07.frx":16FC
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Búsqueda de Cuenta por Código o Descripcion"
         Top             =   3660
         Width           =   750
      End
      Begin VB.CommandButton Command5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   120
         Picture         =   "FPLANCUEN07.frx":2A62
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Ordenamiento Creciente por Código de Cuenta"
         Top             =   4665
         Width           =   750
      End
      Begin VB.CommandButton Command7 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   120
         Picture         =   "FPLANCUEN07.frx":2D6C
         TabIndex        =   45
         ToolTipText     =   "Agrega Cuenta al Final del Plan"
         Top             =   2160
         Width           =   750
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   120
         Picture         =   "FPLANCUEN07.frx":3076
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1365
         Width           =   750
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   120
         Picture         =   "FPLANCUEN07.frx":3380
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   360
         Width           =   750
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   120
         Picture         =   "FPLANCUEN07.frx":34CA
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Aprueba y Graba el Plan de Cuentas"
         Top             =   840
         Width           =   750
      End
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   705
         TabIndex        =   40
         Top             =   5400
         Width           =   740
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   41
            Top             =   0
            Width           =   12000
         End
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
Attribute VB_Name = "FPLANCUEN07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODOEDIT%, ACTIVLIST$, ILISTACT&
Dim MODIFIC%

Sub ActualizarPlancuenMineparts()
'     If EmpresaNoRealizaCambios% > 0 Then
'        Exit Sub
'     End If
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
        SQLX$ = "SELECT * FROM PLANCUEN WITH(NOLOCK)"
        Set PLANCCCUEN = READSET(SQLX$)
        With PLANCCCUEN
          Do While Not .EOF
            CUEIN% = XVALO(!CODINCUE)
            CONDI$ = " CODINCUE = " & CUEIN%
            Call DuplicarRegistroDeUnaBaseAOtra("AHP", "PLANCUEN", "MINEPARTS", "PLANCUEN", CONDI$, 1)
            .MoveNext
          Loop
        End With
     End If
     'refresco la conexion
     FLEXCONN.Close
     Call ABRECONEXION
     Call TransferirArchivos("AHP", "MINEPARTS")

     'esta pendiente copiar los dat para actualizar automaticamente en MINEPARTS
     'VER COMO ENCONTRAR LA RUTA DE MINEPARTS.
     'Call FileCopy(LUDAT$ + ECUECOT + ".DAT", LDX1$ + ARAP$ + ".DAT")


End Sub

'
Private Sub Command1_Click()
   '
   If DERACCE%("MOPLACUE", "Modificar Plan General de Cuentas") < 2 Then
      Exit Sub
   End If
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If EmpresaNoRealizaCambios% > 0 Then
      Exit Sub
   End If
   
   Call CANCELEDIT
   '
   If CONTROL$("PROPSCG", "PLAGRAL") = "GUEST" Then
      Call MENSERR(24, "Sistema Distribuido - Invitado en Plan de Cuentas\Imposible Grabar Plan de Cuentas.")
      Unload Me
      Exit Sub
   End If
   '
   Call GUARDAPLAN
   
   MODIFIC% = 0
   Unload Me
   '
   '
End Sub
'
Sub TransferirArchivos(EmprexaOrigen$, EmprexaDestino$)
EPAC$ = TRIM$(UCase$(EMPREAC$))
If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
   On Error Resume Next
   EmprexaOrigen$ = TRIM$(UCase$(EmprexaOrigen$))
   EmprexaDestino$ = TRIM$(UCase$(EmprexaDestino$))
   
   If TRIM$(EmprexaOrigen$) <> "" And TRIM$(EmprexaDestino$) <> "" Then
     If EXISTE%("FLEXMUL.DRV") = 1 Then
         '
         NX% = FILEOPEN("FLEXMUL.DRV", 0, 79)
         Dim FMX As String * 79
         '
         For i% = 1 To LOF(NX%) / 79
           Get #NX%, i%, FMX$
           SUCAR% = 0
           For j% = 1 To 77
             SUCAR% = SUCAR% + Asc(Mid$(FMX$, j%, 1))
           Next j%
           If CVI(Mid$(FMX$, 78, 2)) = SUCAR% Then
             RASO$ = DESENCRI$(Left$(FMX$, 30))
             PACCE$ = DESENCRI$(Mid$(FMX$, 32, 46))
             If EXISTE%(TRIM$(PACCE$) + "\.") > 0 Then
               RASO$ = TRIM$(UCase$(RASO$))
               If InStr(RASO$, EmprexaOrigen$) > 0 And InStr(RASO$, EmprexaDestino$) < 1 Then
                  RutaOrigen$ = TRIM$(PACCE$) + "\"
               End If
               If InStr(RASO$, EmprexaDestino$) > 0 Then
                  RutaDestinox$ = TRIM$(PACCE$) + "\"
               End If
             End If
           End If
         Next i%
         '
         Close #NX%
     End If
   End If
   If RutaOrigen$ <> "" And RutaDestinox$ <> "" Then
    FileCopy RutaOrigen$ & "CUECON.DAT", RutaDestinox$ & "CUECON.DAT"
    FileCopy RutaOrigen$ & "CUECOT.DAT", RutaDestinox$ & "CUECOT.DAT"
    FileCopy RutaOrigen$ & "CUENTAS.DAT", RutaDestinox$ & "CUENTAS.DAT"
    FileCopy RutaOrigen$ & "DENOCUE.DAT", RutaDestinox$ & "DENOCUE.DAT"
    
    FileCopy RutaOrigen$ & "ECUECON.DAT", RutaDestinox$ & "ECUECON.DAT"
    FileCopy RutaOrigen$ & "ECUECOT.DAT", RutaDestinox$ & "ECUECOT.DAT"
    FileCopy RutaOrigen$ & "INVECUE.DAT", RutaDestinox$ & "INVECUE.DAT"
    FileCopy RutaOrigen$ & "PLAGRAL.DAT", RutaDestinox$ & "PLAGRAL.DAT"
   End If
End If

   On Error GoTo 0
End Sub
Private Sub Command11_Click()
   '
   Command11.Enabled = False
   Call AGRECUENTA(1)
   MODIFIC% = 1
   Command11.Enabled = True
   '
End Sub

Private Sub Command12_Click()
   '
   Command11.Enabled = False
   Call AGRECUENTA(2)
   MODIFIC% = 1
   Command11.Enabled = True
   '
End Sub

Private Sub Command21_Click()
   Command21.Enabled = False
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If EmpresaNoRealizaCambios% > 0 Then
      GoTo 99
   End If
   '
   Call CANCELEDIT
   '
   If TRIM$(Text11) = "" Then
     Call NUEVACUECON
   End If
99 Command21.Enabled = True
End Sub

Sub CANCELEDIT()
   '
   If MODOEDIT% > 0 Then
     '
     LIST1.AddItem ACTIVLIST$, ILISTACT&
     On Error Resume Next
     LIST1.ListIndex = ILISTACT&
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
       U& = LIST1.ListIndex
       If U& < 0 Then GoTo 99
       If U& > LIST1.ListCount - 1 Then GoTo 99
       '
       ZZ$ = LIST1.List(U&)
       LIST1.RemoveItem U&
       Call REPLA(ZZ$, "Baja", 80, 4)
       LIST1.AddItem ZZ$, U&
       MODIFIC% = 1
       On Error Resume Next
       LIST1.ListIndex = U&
       On Error GoTo 0
   End If
99 Command24.Enabled = True
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
   For U& = 1 To LIST1.ListCount
     TTXX$ = LIST1.List(U& - 1)
     TTYY$ = AJUSTI$(Mid$(TTXX$, 9, 12), 12) + AJUSTI$(Mid$(TTXX$, 23, 28), 28)
     JJ& = JJ& + 1
     Call GRAREG("!ECUECOT", TTYY$, JJ&)
   Next U&
   Call SETULTREG("!ECUECOT", JJ&)
   Screen.MousePointer = 1
   Call SELECHO("!ECUECOT")
   '
   ABUS$ = AJUSTI$(VALACT1$("!ECUECOT"), 12)
   For U& = 1 To LIST1.ListCount
     If Mid$(LIST1.List(U& - 1), 9, 12) = ABUS$ Then
       LIST1.ListIndex = U& - 1
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
       U& = LIST1.ListIndex
       If U& < 0 Then Exit Sub
       If U& > LIST1.ListCount - 1 Then Exit Sub
       '
       ZZ$ = LIST1.List(U&)
       LIST1.RemoveItem U&
       Call REPLA(ZZ$, "Baja", 80, 4)
       LIST1.AddItem ZZ$, U&
       MODIFIC% = 1
       On Error Resume Next
       LIST1.ListIndex = U&
       On Error GoTo 0
   End If
   Command4.Enabled = True
   '
End Sub

Private Sub Command5_Click()
   '
   Command5.Enabled = False
   '
   '\\\OT-5-549-FG-29/07/05
   If DERACCE%("ORDEPLAN", "Ordenamiento del Plan de Cuentas") < 1 Then
      GoTo 99
   End If
   If COMALTER%("Este Proceso Re-ordena el Plan de Cuentas en forma Masiva por Número de Cuenta.\Está Seguro de Continuar ?\\No, Cancelar\Si, Continuar") = 1 Then
     GoTo 99    ' NO ORDENA EL PLAN DE CUENTAS
   End If
   '\\\OT-5-549-FG-FIN
   '
   Screen.MousePointer = 11
   FIN& = LIST1.ListCount
   For i& = 1 To FIN&
     Call TRACE(20, i&, FIN&)
     TTI$ = LIST1.List(i& - 1)
     For j& = i& + 1 To FIN&
       TTJ$ = LIST1.List(j& - 1)
       If Mid$(TTI$, 9, 12) > Mid$(TTJ$, 9, 12) Then
         TTZ$ = TTI$
         TTI$ = TTJ$
         TTJ$ = TTZ$
         LIST1.RemoveItem (i& - 1)
         LIST1.AddItem TTI$, i& - 1
         LIST1.RemoveItem (j& - 1)
         LIST1.AddItem TTJ$, j& - 1
       End If
     Next j&
   Next i&
   Screen.MousePointer = 1
   '
99 Command5.Enabled = True
   '
End Sub

Sub Command6_Click()
   Command6.Enabled = False
   Call CANCELEDIT
   Call FINAL("*LICUECO")
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    If DERACCE%("IMPLANCU", "Importacion de Plan de Cuentas Contables") >= 2 Then
       If COMALTER%("¿ Está Seguro de Re-Importar\el Plan de Cuentas Contables ?\  \(Esta Transacción Borrará toda Información Anterior)\\No, Conservar Plan Actual\Si, Re-Importar") = 2 Then
          IMPLANCUE09.Show 1
          Call CARGAPLAN
       End If
    End If
    Command7.Enabled = True
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
    Call APLICACIONSKINS(Me, Picture2)
    Call CENTRAFORM(Me)
    DoEvents
    '
    If TRIM$(EJACT$) <> "" Then
      Caption = Caption + " - " + TRIM$(EJACT$)
    End If
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Call CARGAPLAN
    '
    Call AplicaFuenteAnchoFijo(LIST1, , 10)
    '
    
End Sub
'

Private Sub Form_Unload(Cancel As Integer)
    If MODIFIC% > 0 Then
      If COMALTER%("Se ha Modificado el Plan General de Cuentas\\Registrar Cambios\Descartarlos") = 1 Then
        Call Command1_Click    ' guardar plan de cuentas
      End If
    End If
    '
    TIASI% = 0
    Call CIERRACONTA
End Sub

Private Sub LIST1_Click()
   If MODOEDIT% < 1 Then
     CUEXX$ = Mid$(LIST1.TEXT, 9, 12)
     CUEIIX% = CUEINT(CUEXX$)
     Text14 = OBSERCUE$(CUEIIX%)
   End If
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
    If MODIFIC% > 0 Then
      If COMALTER%("Se ha Modificado el Plan General de Cuentas\\Registrar Cambios\Descartarlos") = 1 Then
        Call Command1_Click    ' guardar plan de cuentas
      End If
    End If
    '
    Call CIERRARCH("*.*")
    Unload Me
    '
End Sub


Private Sub Command31_Click()
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
    LIST1.Clear
    JJ& = 0
    '
    UCUEPLAN$ = "": UREPLA% = 0
    FIN& = LOF(NXP%) / 4
    For kk% = 1 To LOF(NXP%) / 4
       UU& = kk%
       Call TRACE(20, UU&, FIN&)
       Get #NXP%, kk%, IVP$
       CUE% = CVI(Mid$(IVP$, 1, 2))
       NIVEI% = CVI(Mid$(IVP$, 3, 2))
       If NIVEI% < 0 Then NIVEI% = 0
       If NIVEI% > 5 Then NIVEI% = 5
       '
       If CUE% > 0 Then
         ZZ$ = Space$(64)
         If NIVEI% >= 0 Then
           If NIVEI% <= 4 Then
             NVX$ = Left$(String$(NIVEI%, ".") + TRIM$(Str$(NIVEI%)) + "......", 6)
             Call REPLA(ZZ$, NVX$, 1, 6)
             Call REPLA(ZZ$, CODCUE$(CUE%), 10, 12)
             Call REPLA(ZZ$, DENOCUE$(CUE%), 26, 32)
             LIST1.AddItem ZZ$
             UCUEPLAN$ = Mid$(ZZ$, 10, 12)
             If YAVERI% < 1 Then YAPLA%(CUE%) = 1
             UREPLA% = kk%
           End If
         End If
       End If
       '
    Next kk%
    '
    If YAVERI% < 1 Then
      FIN& = LOF(NXQ%) / 4
      For kk% = 1 To LOF(NXQ%) / 4
        UU& = kk%
        Call TRACE(20, UU&, FIN&)
        Get #NXQ%, kk%, IVQ$
        CUE% = CVI(Mid$(IVQ$, 1, 2))
        If CUE% > 0 Then
          If YAPLA%(CUE%) < 1 Then
            NIVEI% = CVI(Mid$(IVQ$, 3, 2))
            If NIVEI% < 0 Then NIVEI% = 0
            If NIVEI% > 5 Then NIVEI% = 5
            '
            COCUE$ = CODCUE$(CUE%)
            If COCUE$ > UCUEPLAN$ Then
              ZZ$ = Space$(64)
              If NIVEI% >= 0 Then
                If NIVEI% <= 5 Then
                  NVX$ = Left$(String$(NIVEI%, ".") + TRIM$(Str$(NIVEI%)) + "......", 6)
                  Call REPLA(ZZ$, NVX$, 1, 6)
                  Call REPLA(ZZ$, CODCUE$(CUE%), 10, 12)
                  Call REPLA(ZZ$, DENOCUE$(CUE%), 26, 32)
                  LIST1.AddItem ZZ$
                  YAPLA%(CUE%) = 1
                End If
              End If
            End If
          End If
        End If
      Next kk%
      YAVERI% = 1
    End If
    '
    LIST1.Refresh
    Screen.MousePointer = 1
    '
End Sub
'
Sub GUARDAPLAN()
   '
   If EmpresaNoRealizaCambios% > 0 Then
      Exit Sub
   End If
   
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   
   Call COPYSTRU("CUECON", "CUECOT")
   Call BLOQARCH("CUENTAS")
   Call BLOQARCH("DENOCUE")
   Call BLOQARCH("INVECUE")
   Call BLOQARCH("PLAGRAL")
   Call BLOQARCH("PLANCUEN")
   Call BLOQARCH("CUECON")
   Call BLOQARCH("CUECOT")
   Call BLOQARCH("ECUECON")
   Call BLOQARCH("ECUECOT")
   '
   JJ& = 0: J4& = 0: J5& = 0: J6& = 0: J7& = 0
   Screen.MousePointer = 11
   '
20 On Error Resume Next
   Plancuen.Close
   On Error GoTo 0
   '
   On Error Resume Next
   'Set Plancuen = Contable.OpenRecordset("PLANCUEN", dbOpenDynaset)
   SQLX$ = "SELECT * FROM PLANCUEN "
   Dim PLANNNCUEN As ADODB.Recordset
   Set PLANNNCUEN = FLEXCONN.Execute(FILTSQL$(SQLX$))

   If Err Then
     If COMALTER%("No se Puede Abrir 'PLANCUEN'\en Modo Exclusivo.\\Reintentar\Cancelar") <> 2 Then GoTo 20
     GoTo 90
   End If
   '
   With PLANNNCUEN
    On Error Resume Next
    'Plancuen.MoveLast
    .MoveLast
    'NUREPLAN% = PLANNNCUEN.RecordCount
    NUREPLAN% = .RecordCount 'NO ESTA FUNCIONANDO ,PERO NO SE USA 'CONSULTAR'
    'Plancuen.MoveFirst
    .MoveFirst
    On Error GoTo 0
   End With
   Set PLANNNCUEN = Nothing

  '
   FIN& = LIST1.ListCount
   URECU& = 0
   For U& = 1 To FIN&
     ZZ$ = LIST1.List(U& - 1)
     CUEIN% = Val(Mid$(ZZ$, 121, 8))
     If CUEIN% > URECU& Then URECU& = CUEIN%
   Next U&
   '
   List2.Clear
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     AGRECUE% = 0
     ZZ$ = LIST1.List(U& - 1)
     COCUE$ = TRIM$(Mid$(ZZ$, 9, 12))
     DECUE$ = AJUSTI$(Mid$(ZZ$, 129, 128), 100)
     CUEIN% = Val(Mid$(ZZ$, 121, 8))
       If CUEIN% < 1 Or CUEIN% > URECU& Then
         URECU& = URECU& + 1
         CUEIN% = URECU&
         AGRECUE% = 1
       End If
     RECUENTA& = CUEIN%
     TICU$ = Mid$(ZZ$, 69, 5)
     IBLE% = 1: If UCase$(Mid$(ZZ$, 76, 1)) = "N" Then IBLE% = 0
     STAT% = 1: If UCase$(Mid$(ZZ$, 80, 1)) = "B" Then STAT% = 9
     NIVEI% = Val(TRIM$(REPLACAR$(TRIM$(Left$(ZZ$, 6)), ".", " ")))
     '
     SQLX$ = "SELECT * FROM PLANCUEN "
     SQLX$ = SQLX$ + " WHERE CODINCUE = " & CUEIN% & ""
     Dim PLANCCCUEN As ADODB.Recordset
     Set PLANCCCUEN = New ADODB.Recordset
     PLANCCCUEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

     With PLANCCCUEN
       On Error Resume Next
       '.FindFirst "CODINCUE = " & CUEIN%
       'If .NoMatch = True Then
       .MoveFirst
       If Err > 0 Then
           .AddNew   ' YA CONVERTIDO
           On Error Resume Next
           !CodiEmpr = CodiEmp%
           On Error GoTo 0
           On Error GoTo 0
         'Else
           '.Edit
       End If
       On Error GoTo 0
       '
       !CodExCue = COCUE$
       !DENOCUE = DECUE$
       !CODINCUE = CUEIN%
       !STATCUE = STAT%
       !NIVAGRU = NIVEI%
       !NUORDIT = U&
       !TipoCue = TICU$
       !POSIMPU = IBLE%
       '
       .Update
       '.Bookmark = .LastModified
       '.MoveNext
       '
     End With
     '
     XCUE$ = String$(4, 0)
     Call REPLA(XCUE$, MKI$(CUEIN%), 1, 2)
     Call REPLA(XCUE$, MKI$(NUVEI%), 3, 2)
     Call GRAREG("PLAGRAL", XCUE$, U&)
     '
   Next U&
   Set PLANCCCUEN = Nothing
   'Plancuen.Close
   Call SETULTREG("PLAGRAL", FIN&)
   Call CIERRARCH("PLAGRAL")
   '
   '
   SQLX$ = "SELECT * FROM PLANCUEN"
   SQLX$ = SQLX$ + " ORDER BY CODEXCUE ASC"
   'Set MASTERECO = Contable.OpenRecordset("SELECT * FROM PLANCUEN ORDER BY CODEXCUE ASC;", dbOpenDynaset, dbReadOnly)
   Dim MASTTTERECO As ADODB.Recordset
   Set MASTTTERECO = New ADODB.Recordset
   MASTTTERECO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   With MASTTTERECO
      On Error Resume Next
      'MASTERECO.MoveLast
      .MoveLast
      If Err Then On Error GoTo 0: GoTo 90
      On Error GoTo 0
      'CAREC% = MASTERECO.RecordCount
      CAREC% = .RecordCount
      If CAREC% < 1 Then GoTo 90
      '
      FIN& = CAREC%
      REBLA$ = REGBLAN$("ECUECON")
      J1& = 0: J2& = 0: J3& = 0: J4& = 0
      .MoveFirst
      '
      Do While Not .EOF
        Call TRACE(20, J1&, FIN&)
        EL1$ = REBLA$
        CUEXINT% = !CODINCUE
        
        '
        If CUEXINT% > 0 Then
          While ULTREG&("CUENTAS") < CUEXINT%
            Call REGAPP("CUENTAS", Space$(13) + String$(50, 0) + Chr$(1))
          Wend
          While ULTREG&("DENOCUE") < CUEXINT%
            Call REGAPP("DENOCUE", "." + Space$(127))
          Wend
          RECUE& = CUEXINT%
          '
          XCUE$ = REGLEIDO$("CUENTAS", RECUE&)
          Call REPLA(XCUE$, !CodExCue, 1, 12)
          Call REPLA(XCUE$, !TipoCue, 13, 1)
          Call GRAREG("CUENTAS", XCUE$, RECUE&)
          '
          XCUE$ = REGLEIDO$("DENOCUE", RECUE&)
          Call REPLA(XCUE$, !DENOCUE, 1, 100)
          Call REPLA(XCUE$, !DENOCUE, 101, 28)
          Call GRAREG("DENOCUE", XCUE$, RECUE&)
          '
          XCUE$ = REGBLAN$("INVECUE")
          Call REPLA(XCUE$, !CodExCue, 1, 12)
          Call REPLA(XCUE$, MKI$(CUEXINT%), 13, 2)
          J4& = J4& + 1
          Call GRAREG("INVECUE", XCUE$, J4&)
          '
        End If
        '
        Call REPLA(EL1$, !CodExCue, 1, 12)
        Call REPLA(EL1$, !DENOCUE, 13, 28)
        Call REPLA(EL1$, MKI$(CUEXINT%), 41, 2)
        J1& = J1& + 1
        Call GRAREG("ECUECOT", EL1$, J1&)
        '
        If !STATCUE = 1 Then
          If !POSIMPU = 1 Then
            J2& = J2& + 1
            Call GRAREG("ECUECON", EL1$, J2&)
            '
            Z4$ = REGBLAN$("CUECON")
            Call REPLA(Z4$, MKI$(CUEXINT%), 1, 2)
            Call REPLA(Z4$, !DENOCUE, 3, 30)
            Call REPLA(Z4$, !CodExCue, 33, 12)
            Call REPLA(Z4$, !CodExCue, 45, 12)
            J7& = J7& + 1
            Call GRAREG("CUECOT", Z4$, J7&)
            '
            TICU$ = UCase$(Left$(!TipoCue, 1))
            If TICU$ = "I" Or TICU$ = "G" Then
              J3& = J3& + 1
              Call GRAREG("CUECON", Z4$, J3&)
            End If
          End If
        End If
        '
        .MoveNext
      Loop
   '
   End With
   '
   MASTTTERECO.Close
   Set MASTTTERECO = Nothing
   '

   
   Call SETULTREG("ECUECOT", J1&)
   Call SETULTREG("ECUECON", J2&)
   Call SETULTREG("CUECON", J3&)
   Call SETULTREG("CUECOT", J7&)
   Call SETULTREG("INVECUE", J4&)
   Call CIERRARCH("ECUECOT"): Call FRESHECHO("ECUECOT")
   Call CIERRARCH("ECUECON"): Call FRESHECHO("ECUECON")
   Call CIERRARCH("CUECON"): Call FRESHECHO("CUECON")
   Call CIERRARCH("CUECOT"): Call FRESHECHO("CUECOT")
   Call CIERRARCH("CUENTAS"): Call FRESHECHO("CUENTAS")
   Call CIERRARCH("DENOCUE")
   Call CIERRARCH("INVECUE")
   Call CIERRARCH("PLAGRAL"): Call FRESHECHO("PLAGRAL")
   Call CIERRARCH("PLANCUEN"): Call FRESHECHO("PLANCUEN")
   ' LOS FRESHECHO SON PARA REFRESCAR LAS LISTAS DE SELECCION
   '
   ABCT$ = LOADFILE$(LUDAT$ + "CUENTAS.DAT"): Call SAVEFILE(LUDAT$ + "CUENTAS.OLD", ABCT$)
   ABCT$ = LOADFILE$(LUDAT$ + "DENOCUE.DAT"): Call SAVEFILE(LUDAT$ + "DENOCUE.OLD", ABCT$)
   ABCT$ = LOADFILE$(LUDAT$ + "INVECUE.DAT"): Call SAVEFILE(LUDAT$ + "INVECUE.OLD", ABCT$)
   ABCT$ = LOADFILE$(LUDAT$ + "PLAGRAL.DAT"): Call SAVEFILE(LUDAT$ + "PLAGRAL.OLD", ABCT$)
   '
   Call ActualizarPlancuenMineparts

90 'Plancuen.Close
   Screen.MousePointer = 1
   '
End Sub
Sub CARGAPLAN()
   '
   Screen.MousePointer = 11
   LIST1.Clear
   '
   Call APERBASDAT("PLANCUEN")
   Dim Plansort As ADODB.Recordset
   Set Plansort = New ADODB.Recordset
   strsql = "SELECT * FROM PLANCUEN ORDER BY NuOrdIt ASC"
   Set Plansort = FLEXCONN.Execute(FILTSQL$(strsql))       ', FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   U& = 0
   With Plansort
     Do While Not .EOF
       U& = 1
       ZZ$ = Space$(256)
       '
10     COCUE$ = SAFETEXT$(!CodExCue)
       DECUE$ = SAFETEXT$(!DENOCUE)
       NIVEI% = XVALO(!NIVAGRU)
         NVX$ = Left$(String$(NIVEI%, ".") + TRIM$(Str$(NIVEI%)) + "......", 6)
         '
       Call REPLA(ZZ$, NVX$, 1, 6)
       Call REPLA(ZZ$, COCUE$, 9, 12)
       Call REPLA(ZZ$, DECUE$, 23, 44)
       Call REPLA(ZZ$, DECUE$, 129, 128)
       '
       TICUEX$ = TRIM$(UCase$(Left$(SAFETEXT$(!TipoCue), 1)))
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
       If XVALO(!POSIMPU) < 1 Then IMPUBLE$ = "No"
       Call REPLA(ZZ$, IMPUBLE$, 76, 2)
       '
       STATUX$ = ""
       If XVALO(!STATCUE) = 9 Then STATUX$ = "Baja"
       Call REPLA(ZZ$, STATUX$, 80, 4)
       '
       NUINT$ = AJUSTD$(Str$(XVALO(!CODINCUE)), 8)
       Call REPLA(ZZ$, NUINT$, 121, 8)
       '
       LIST1.AddItem ZZ$
       '
     .MoveNext
     Loop
     '
   End With
   '
90 LIST1.Refresh
   Screen.MousePointer = 1
   '
End Sub
'
Sub CARPLANCUEN()   ' CARGA DE PLANCUEN
   '
   LIST1.Clear
   If ULTREG("PLANCUEN") < 1 Then
     Call CARPLAGRAL ' carga de PLAGRAL.DAT
     GoTo 90
   End If
   '
   FIN& = ULTREG&("PLANCUEN")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("PLANCUEN", U&)
     ZZ$ = Space$(256)
     '
     COCUE$ = Left$(XX$, 12)
     DECUE$ = Mid$(XX$, 129, 128)
     NIVEI% = Asc(Mid$(XX$, 62, 1))
     NVX$ = Left$(String$(NIVEI%, ".") + TRIM$(Str$(NIVEI%)) + "......", 6)
     '
     Call REPLA(ZZ$, NVX$, 1, 6)
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
     LIST1.AddItem ZZ$
     '
   Next U&
   '
90 LIST1.Refresh
   '
End Sub
'
Sub CARPLAGRAL()
   '
   FIN& = ULTREG&("PLAGRAL")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     IVP$ = REGLEIDO$("PLAGRAL", U&)
     '
     CUE% = CVI(Mid$(IVP$, 1, 2))
     NIVEI% = CVI(Mid$(IVP$, 3, 2))
     If NIVEI% < 0 Then NIVEI% = 0
     If NIVEI% > 5 Then NIVEI% = 5
     '
     If CUE% > 0 Then
       ZZ$ = Space$(256)
       If NIVEI% >= 0 Then
         If NIVEI% <= 5 Then
           NVX$ = Left$(String$(NIVEI%, ".") + TRIM$(Str$(NIVEI%)) + "......", 6)
           COCUE$ = CODCUE$(CUE%)
           DECUE$ = DENOCUE$(CUE%)
           '
           Call REPLA(ZZ$, NVX$, 1, 6)
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
           LIST1.AddItem ZZ$
           '
         End If
       End If
     End If
   Next U&
   '
   If ULTREG&("PLAGRAL") < 1 Then
     FIN& = ULTREG&("ECUECOT")
     If FIN& > 0 Then
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         IVP$ = REGLEIDO$("ECUECOT", U&)
         '
         CUE% = CVI(Mid$(IVP$, 41, 2))
         NIVEI% = 2
         '
         If CUE% > 0 Then
           ZZ$ = Space$(256)
           NVX$ = Left$(String$(NIVEI%, ".") + TRIM$(Str$(NIVEI%)) + "......", 6)
           COCUE$ = TRIM$(Left$(IVP$, 12))
           DECUE$ = TRIM$(Mid$(IVP$, 13, 28))
           '
           If COCUE$ <> "" Then
             Call REPLA(ZZ$, NVX$, 1, 6)
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
             LIST1.AddItem ZZ$
             '
           End If
         End If
       Next U&
     End If
   End If
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
   Label2 = 5 - VScroll1.Value
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
     FIN& = LIST1.ListCount
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       If UCase$(TRIM$(Mid$(LIST1.List(U& - 1), 9, 12))) = UCase$(CONUCUE$) Then
         Call MENSERR(24, "Cuenta Registrada en el Plan de Cuentas")
         LIST1.ListIndex = U& - 1
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
     If LIST1.ListIndex <= 0 Then
         NVLX% = 0
       Else
         NVLX% = Val(TRIM$(REPLACAR$(TRIM$(Left$(LIST1.TEXT, 5)), ".", " ")))
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
   FIN& = LIST1.ListCount
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     If UCase$(TRIM$(Mid$(LIST1.List(U& - 1), 9, 12))) = UCase$(COCUE$) Then
       Call MENSERR(24, "Cuenta Existente en el Plan de Cuentas")
       LIST1.ListIndex = U& - 1
       Exit Sub
     End If
   Next U&
   '
   NIVEI% = Val(Label2)
   If NIVEI% < 0 Or NIVEI% > 5 Then
      Call MENSERR(24, "Mal Nivel de Agrupación")
      Exit Sub
   End If
   '
   ' aqui validar el nivel de agrupacion segun cuenta anterior
   '
   ZZ$ = Space$(256)
   NVX$ = Left$(String$(NIVEI%, ".") + TRIM$(Str$(NIVEI%)) + "......", 6)
   Call REPLA(ZZ$, NVX$, 1, 6)
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
   If LIST1.ListIndex < 0 Then MODOI% = 1   ' SI NO DA ERROR
   '
   If MODOI% = 1 Then
       LIST1.AddItem ZZ$
       On Error Resume Next
       LIST1.ListIndex = LIST1.ListCount - 1
       On Error GoTo 0
     Else
       ILISTA& = LIST1.ListIndex
       LIST1.AddItem ZZ$, ILISTA&
       On Error Resume Next
       LIST1.ListIndex = ILISTA&
       On Error GoTo 0
   End If
   '
   CUEIXI% = Val(Label8)
   If CUEIXI% > 0 Then
     SQLX$ = "SELECT * FROM PLANCUEN "
     SQLX$ = SQLX$ + " WHERE CodInCue = " & CUEIXI% & ""
     'Set Plancuen = Contable.OpenRecordset("PLANCUEN", dbOpenDynaset)
     Dim PLANNNCUEN As ADODB.Recordset
     Set PLANNNCUEN = New ADODB.Recordset
     PLANNNCUEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     With PLANNNCUEN
      On Error Resume Next
      .MoveFirst
      If Err Then GoTo 18
'     With Plancuen
       '.FindFirst "CodInCue = " & CUEIXI%
       'If .NoMatch = False Then
         '.Edit
         !Observa = Text14
         .Update
         GoTo 19
       'End If
     End With
   End If
   Set PLANNNCUEN = Nothing
   If TRIM$(Text14) <> "" Then
18   On Error GoTo 0
     Call MENSERR(24, "Imposible Actualizar Observaciones\en el Contexto Actual.\  \Vuelva a Intentarlo en Otro Momento.")
   End If
   '
19 MODOEDIT% = 0
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
   On Error Resume Next
   Call LIST1_Click
   On Error GoTo 0
   '
End Sub

Sub EDITACUEN()
   '
   If MODOEDIT% > 0 Then
     Call MENSERR(24, "Primero Guarde la Cuenta Editada")
     Exit Sub
   End If
   '
   U& = LIST1.ListIndex
   If U& < 0 Then Exit Sub
   If U& > LIST1.ListCount - 1 Then Exit Sub
   '
   MODOEDIT% = 1
   ILISTACT& = U&
   ACTIVLIST$ = LIST1.List(U&)
   Command4.Enabled = False
   Command8.Enabled = False
   Command24.Enabled = False
   '
   URECU& = ULTREG&("CUENTAS")
   ZZ$ = LIST1.List(U&)
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
   NIVEI% = Val(TRIM$(REPLACAR$(TRIM$(Left$(ZZ$, 6)), ".", " ")))
   Label2 = TRIM$(Str$(NIVEI%))
   VScroll1.Value = 5 - NIVEI%
   '
   LIST1.RemoveItem U&
   If U& < LIST1.ListCount Then
      LIST1.ListIndex = U&
   End If
   '
   Text14 = OBSERCUE$(CUEIN%)
   '
End Sub
