VERSION 5.00
Begin VB.Form OPMASCL04 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maestro de Clientes - Configuración"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11130
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   255
      Left            =   8760
      TabIndex        =   72
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   8350
      TabIndex        =   66
      Top             =   420
      Width           =   330
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   7085
      TabIndex        =   65
      Top             =   315
      Width           =   330
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   7575
      ScaleHeight     =   75
      ScaleWidth      =   540
      TabIndex        =   63
      Top             =   315
      Width           =   600
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   5000
         Left            =   -420
         TabIndex        =   64
         Top             =   -105
         Width           =   12000
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Control"
      Height          =   950
      Left            =   9135
      TabIndex        =   60
      Top             =   240
      Width           =   1800
      Begin VB.CommandButton command1 
         Height          =   540
         Left            =   945
         Picture         =   "OPMASCL04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   280
         Width           =   645
      End
      Begin VB.CommandButton Command2 
         Height          =   540
         Left            =   220
         Picture         =   "OPMASCL04.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   280
         Width           =   645
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "I.B.Pcia.Bs.As"
      Height          =   950
      Left            =   4935
      TabIndex        =   57
      Top             =   240
      Width           =   1800
      Begin VB.CommandButton Command3 
         Height          =   540
         Left            =   210
         Picture         =   "OPMASCL04.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Configuracion Funcionamiento Percepciones IB Pcia de Bs.As"
         Top             =   280
         Width           =   645
      End
      Begin VB.CommandButton Command4 
         Height          =   540
         Left            =   945
         Picture         =   "OPMASCL04.frx":0896
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Listado Clientes para I.Brutos"
         Top             =   280
         Width           =   645
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Valores Preteterminados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3880
      Left            =   4935
      TabIndex        =   24
      Top             =   1365
      Width           =   6000
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   3375
         Left            =   60
         TabIndex        =   28
         Top             =   370
         Width           =   5840
         Begin VB.CommandButton BOTSEL 
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
            Height          =   300
            Index           =   0
            Left            =   1890
            TabIndex        =   49
            Top             =   840
            Width           =   175
         End
         Begin VB.TextBox COTEXT 
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
            Index           =   0
            Left            =   1365
            TabIndex        =   48
            Text            =   "  "
            Top             =   840
            Width           =   570
         End
         Begin VB.TextBox DETEXT 
            BackColor       =   &H00E0E0E0&
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
            Index           =   0
            Left            =   2205
            TabIndex        =   47
            TabStop         =   0   'False
            Text            =   "  "
            Top             =   840
            Width           =   3560
         End
         Begin VB.TextBox DETEXT 
            BackColor       =   &H00E0E0E0&
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
            Index           =   4
            Left            =   2205
            TabIndex        =   46
            TabStop         =   0   'False
            Text            =   "  "
            Top             =   1680
            Width           =   3560
         End
         Begin VB.TextBox DETEXT 
            BackColor       =   &H00E0E0E0&
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
            Index           =   3
            Left            =   2205
            TabIndex        =   45
            TabStop         =   0   'False
            Text            =   "  "
            Top             =   1260
            Width           =   3560
         End
         Begin VB.TextBox COTEXT 
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
            Index           =   4
            Left            =   1365
            TabIndex        =   44
            Text            =   " "
            Top             =   1680
            Width           =   570
         End
         Begin VB.TextBox COTEXT 
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
            Index           =   2
            Left            =   1365
            TabIndex        =   43
            Text            =   " "
            Top             =   420
            Width           =   570
         End
         Begin VB.TextBox DETEXT 
            BackColor       =   &H00E0E0E0&
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
            Index           =   2
            Left            =   2205
            TabIndex        =   42
            TabStop         =   0   'False
            Text            =   "  "
            Top             =   420
            Width           =   3560
         End
         Begin VB.TextBox COTEXT 
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
            Index           =   3
            Left            =   1365
            TabIndex        =   41
            Text            =   " "
            Top             =   1260
            Width           =   570
         End
         Begin VB.TextBox COTEXT 
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
            Index           =   5
            Left            =   1365
            TabIndex        =   40
            Text            =   " "
            Top             =   2100
            Width           =   570
         End
         Begin VB.TextBox DETEXT 
            BackColor       =   &H00E0E0E0&
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
            Index           =   5
            Left            =   2205
            TabIndex        =   39
            TabStop         =   0   'False
            Text            =   "  "
            Top             =   2100
            Width           =   3560
         End
         Begin VB.CommandButton BOTSEL 
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
            Height          =   300
            Index           =   2
            Left            =   1890
            TabIndex        =   38
            Top             =   420
            Width           =   175
         End
         Begin VB.CommandButton BOTSEL 
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
            Height          =   300
            Index           =   3
            Left            =   1890
            TabIndex        =   37
            Top             =   1260
            Width           =   175
         End
         Begin VB.CommandButton BOTSEL 
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
            Height          =   300
            Index           =   4
            Left            =   1890
            TabIndex        =   36
            Top             =   1680
            Width           =   175
         End
         Begin VB.CommandButton BOTSEL 
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
            Height          =   300
            Index           =   5
            Left            =   1890
            TabIndex        =   35
            Top             =   2100
            Width           =   175
         End
         Begin VB.CommandButton BOTSEL 
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
            Height          =   300
            Index           =   6
            Left            =   5595
            TabIndex        =   34
            Top             =   2625
            Width           =   175
         End
         Begin VB.TextBox DETEXT 
            BackColor       =   &H00E0E0E0&
            Height          =   300
            Index           =   6
            Left            =   1365
            TabIndex        =   33
            TabStop         =   0   'False
            Text            =   "  "
            Top             =   3000
            Width           =   4410
         End
         Begin VB.TextBox COTEXT 
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
            Index           =   6
            Left            =   3360
            TabIndex        =   32
            Text            =   " "
            Top             =   2625
            Width           =   2250
         End
         Begin VB.CommandButton BOTSEL 
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
            Height          =   300
            Index           =   7
            Left            =   1890
            TabIndex        =   31
            Top             =   0
            Width           =   175
         End
         Begin VB.TextBox DETEXT 
            BackColor       =   &H00E0E0E0&
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
            Index           =   7
            Left            =   2205
            TabIndex        =   30
            TabStop         =   0   'False
            Text            =   "  "
            Top             =   0
            Width           =   3560
         End
         Begin VB.TextBox COTEXT 
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
            Index           =   7
            Left            =   1365
            TabIndex        =   29
            Text            =   " "
            Top             =   0
            Width           =   570
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Provincia:"
            Height          =   225
            Left            =   105
            TabIndex        =   56
            Top             =   945
            Width           =   1185
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor:"
            Height          =   330
            Left            =   -630
            TabIndex        =   55
            Top             =   2205
            Width           =   1905
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cond.Pago:"
            Height          =   225
            Left            =   -420
            TabIndex        =   54
            Top             =   1785
            Width           =   1710
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Categoría I-B"
            Height          =   330
            Left            =   -105
            TabIndex        =   53
            Top             =   525
            Width           =   1380
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo Cliente:"
            Height          =   225
            Left            =   -105
            TabIndex        =   52
            Top             =   105
            Width           =   1380
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L.Precios:"
            Height          =   225
            Left            =   -420
            TabIndex        =   51
            Top             =   1365
            Width           =   1710
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta Madre:"
            Height          =   330
            Left            =   1365
            TabIndex        =   50
            Top             =   2730
            Width           =   1905
         End
      End
   End
   Begin VB.TextBox COTEXT 
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
      Index           =   1
      Left            =   6300
      TabIndex        =   27
      Text            =   " "
      Top             =   2100
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox DETEXT 
      BackColor       =   &H00E0E0E0&
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
      Index           =   1
      Left            =   7140
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   2100
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton BOTSEL 
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
      Height          =   300
      Index           =   1
      Left            =   6825
      TabIndex        =   25
      Top             =   2100
      Visible         =   0   'False
      Width           =   175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Controles Operativos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2245
      Left            =   210
      TabIndex        =   19
      Top             =   3000
      Width           =   4530
      Begin VB.Frame Frame12 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Formato Número de CUIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   240
         TabIndex        =   67
         Top             =   1440
         Width           =   4110
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Suprime Validación de Formato (Argentina)"
            Height          =   255
            Left            =   105
            TabIndex        =   68
            Top             =   300
            Width           =   3960
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Alta de Nueva Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   240
         TabIndex        =   20
         Top             =   300
         Width           =   4110
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   3255
            TabIndex        =   23
            Top             =   615
            Width           =   645
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Primer Número Libre Mayor que ..."
            Height          =   750
            Left            =   1785
            TabIndex        =   22
            Top             =   195
            Width           =   1485
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Número Siguiente al Ultimo"
            Height          =   750
            Left            =   315
            TabIndex        =   21
            Top             =   195
            Value           =   -1  'True
            Width           =   1275
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tablas de Validación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4530
      Begin VB.CommandButton Command20 
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
         Left            =   3255
         TabIndex        =   1
         Top             =   285
         Width           =   175
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Opciones de Trabajo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         Left            =   210
         TabIndex        =   17
         Top             =   1080
         Width           =   4110
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Desc.Combin."
            Height          =   240
            Index           =   11
            Left            =   2310
            TabIndex        =   71
            Top             =   1260
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Zonas de Reparto"
            Height          =   350
            Index           =   10
            Left            =   2310
            TabIndex        =   70
            Top             =   850
            Width           =   1485
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Vendedores"
            Height          =   240
            Index           =   4
            Left            =   105
            TabIndex        =   69
            Top             =   1260
            Width           =   1590
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Grupos Clientes"
            Height          =   240
            Index           =   9
            Left            =   2310
            TabIndex        =   16
            Top             =   570
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Ctas. Contables"
            Height          =   240
            Index           =   8
            Left            =   2310
            TabIndex        =   15
            Top             =   250
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Listas de Precios"
            Height          =   240
            Index           =   7
            Left            =   105
            TabIndex        =   12
            Top             =   250
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Condiciones de Pago"
            Height          =   240
            Index           =   6
            Left            =   105
            TabIndex        =   13
            Top             =   570
            Width           =   2220
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Provincias y Zonas de Venta"
            Height          =   350
            Index           =   5
            Left            =   105
            TabIndex        =   14
            Top             =   850
            Width           =   1800
         End
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Abrir"
         Height          =   700
         Left            =   3600
         Picture         =   "OPMASCL04.frx":0F00
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   285
         Width           =   750
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   " "
         Top             =   705
         Width           =   3225
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1680
         TabIndex        =   9
         Text            =   " "
         Top             =   285
         Width           =   1695
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   2
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command21 
            Caption         =   "Análisis A-B-C"
            Height          =   435
            Left            =   2205
            TabIndex        =   8
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Estadísticas"
            Height          =   435
            Left            =   210
            TabIndex        =   7
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command23 
            Caption         =   "Críticos"
            Height          =   435
            Left            =   2205
            TabIndex        =   6
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command24 
            Caption         =   "Consignación"
            Height          =   435
            Left            =   210
            TabIndex        =   5
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command25 
            Caption         =   "Negativos"
            Height          =   435
            Left            =   2205
            TabIndex        =   4
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command26 
            Caption         =   "Inventarios"
            Height          =   435
            Left            =   210
            TabIndex        =   3
            Top             =   420
            Width           =   1590
         End
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tabla:"
         Height          =   225
         Left            =   -735
         TabIndex        =   18
         Top             =   345
         Width           =   2325
      End
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   7140
      Picture         =   "OPMASCL04.frx":120A
      Top             =   770
      Width           =   1515
   End
End
Attribute VB_Name = "OPMASCL04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TABLA$(16)
Dim MODUCONE$

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Sub Command16_Click()
    '
    Command16.Enabled = False
    If TRIM$(MODUCONE$) <> "" Then
    If TRIM$(MODUCONE$) = "GRUCOCLI" Then
            LDT$ = LUDAT$
            Call CIERRARCH("GRUCOCLI")
            Call CIERRARCH("!GRUCOCLI")
            'Call COPYFILE(LDT$ + "GRUCOCLI.DAT", "C:\FLEXOFT\GRUCOCLI.WKF")
            Call TRALOCAL("GRUCOCLI", "")
            VTB% = VENTABU%("GRUCOCLI")
            If VTB% < 0 Then GoTo 99
            J& = 0
            For I& = 1 To ULTREG&("!GRUCOCLI")
                X$ = REGLEIDO$("!GRUCOCLI", I&)
                If CVI(Left$(X$, 2)) > 0 Then
                    J& = J& + 1
                    Call GRAREG("GRUCOCLI", X$, J&)
                End If
            Next I&
            Call SETULTREG("GRUCOCLI", J&)
            Call CIERRARCH("GRUCOCLI")
          ElseIf MODUCONE = "TADEPOSI" Then
            Call CARTADEPOSI
          ElseIf MODUCONE = "TATIMAT" Then
            Call CARTATIMA
          ElseIf MODUCONE = "CONPAGO" Then
            Call CARCONPAGOS
          ElseIf MODUCONE = "TAPROVIN" Then
            Call CARTAPROVIN
          ElseIf MODUCONE = "CUEINGAS" Then
            Call CARCUEINGAS
          ElseIf MODUCONE = "ZONAREP" Then
            Call CARZONAREP
          ElseIf MODUCONE = "GRUFACLI" Then
            DEFDESCO.Show 1
          ElseIf MODUCONE = "*VENDEDOR" Then
                 RTA% = ALTERNA%("ABM Vendedores\Direccion de E-mail")
                If RTA% = 1 Then
                    Call CONECRUN(MODUCONE$, Text2.TEXT)
                    ElseIf RTA% = 2 Then
                        Call SELECHO("VENDEDOR")
                            OP$ = VALACT1$("VENDEDOR")
                            If OP$ <> "" Then
                                    X$ = FILTERGETRECORD("DIMAIVEN", 1, Chr(OP$))
                                    DIMAI$ = Mid$(X$, 2, 63)
                                    DIMAI$ = InputBox("Modifique o ingrese su E-mail", , DIMAI$)
                                    Call REPLA(X$, Chr$(OP$), 1, 1)
                                    Call REPLA(X$, DIMAI$, 2, 63)
                                    Call FILTERSETRECORD("DIMAIVEN", 1, Chr$(OP$), X$)
                                Else: GoTo 99
                            End If
                    Else: GoTo 99
                End If
            Else
            Call CONECRUN(MODUCONE$, Text2.TEXT)
        End If
    End If
99  Command16.Enabled = True
End Sub

Private Sub Command2_Click()
   OPXX$ = "-1"
   If Option3.Value = True Then OPXX$ = TRIM$(Str$(XVALO(Text3.TEXT)))
   Call GRACONTROL("", "PRONUCLI", OPXX$)
   '\\\OT-05-0198-WAR-19/04/05-///
   VALICUIT1$ = "NO"
   If Check1.Value = 1 Then VALICUIT1$ = "SI"
   Call GRACONTROL("FORMCUIT", "SUPRIVAL", VALICUIT1$)
   ''\\\OT-05-0198-WAR-FIN///
   For Index = 0 To 7
       TV11% = TIPFIELD(ARCHE$(Index), 1)
       LF11% = LENFIELD(ARCHE$(Index), 1)
       ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
       VAPRE$ = "": ECO$ = ""
       If TRIM$(COTEXT(Index).TEXT) <> "" Then
          ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
          If ECO$ <> "" Then
             VAPRE$ = TRIM$(COTEXT(Index).TEXT)
          End If
       End If
       COTEXT(Index).TEXT = VAPRE$
       DETEXT(Index).TEXT = ECO$
       Call GRACONTROL("AMACLI", ARCHE$(Index), VAPRE$)
   Next Index
   '
End Sub

Sub Command3_Click()
   Command3.Enabled = False
   If DERACCE%("ACPAIBRU", "Actualizar Parametros Ingresos Brutos") = 2 Then
10   VDEF$ = String$(16, 0) + Space$(16)
     Call REPLA(VDEF$, MKS$(XVALO(CONTROL$("PERIBRU", "TASA-A"))), 1, 4)
     Call REPLA(VDEF$, MKS$(XVALO(CONTROL$("PERIBRU", "TASA-B"))), 5, 4)
     Call REPLA(VDEF$, MKS$(XVALO(CONTROL$("PERIBRU", "TASA-EXP"))), 9, 4)
     Call REPLA(VDEF$, MKS$(XVALO(CONTROL$("PERIBRU", "MINIPER"))), 13, 4)
     Call REPLA(VDEF$, CONTROL$("PERIBRU", "CUECONTA"), 17, 12)
     '
     OBX$ = OBJPLA$("PARPERIBRU", VDEF$)
     If OBX$ <> "" Then
       Call GRACONTROL$("PERIBRU", "TASA-A", Str(CVS(Mid$(OBX$, 1, 4))))
       Call GRACONTROL$("PERIBRU", "TASA-B", Str(CVS(Mid$(OBX$, 5, 4))))
       Call GRACONTROL$("PERIBRU", "TASA-EXP", Str(CVS(Mid$(OBX$, 9, 4))))
       Call GRACONTROL$("PERIBRU", "MINIPER", Str(CVS(Mid$(OBX$, 13, 4))))
       '\\\OT-06-0087-WAR-13/03/06///
       If TRIM$(Mid$(OBX$, 1, 17)) <> "" Then
          CUECO1$ = Mid$(OBX$, 17, 12)
          If TRIM$(CUECO1$) = "" Or ECOARCH("ECUECON", CUECO1$) = "" Then
            Call MENSERR(24, "Cuenta Contable no Válida.")
            GoTo 10
          End If
       End If
       '\\\OT-06-0087-WAR-FIN///
       Call GRACONTROL$("PERIBRU", "CUECONTA", Mid$(OBX$, 17, 12))
     End If
     '
   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call FILESORT("DAPIBCLI", "DAPIBCLJ", 1, 1, "ASC")
   Screen.MousePointer = 11
   FIN& = ULTREG&("DAPIBCLJ")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("DAPIBCLJ", U&)
     NCLIX% = CVI(Left$(XX$, 2))
     CATIBX% = XVALO(CATIBRU$(NCLIX%))
     NIBR$ = NIBRCLI$(NCLIX%)
     Call REPLA(XX$, MKI$(CATIBX%), 65, 2)
     Call REPLA(XX$, NIBR$, 67, 16)
     If CATIBX% > 0 Then
       If TRIM$(NIBR$) = "" Then
         Call REPLA(XX$, "Falta Nro. Inscripción", 83, 24)
       End If
     End If
     Call GRAREG("DAPIBCLJ", XX$, U&)
   Next U&
   Call CIERRARCH("DAPIBCLJ")
   Screen.MousePointer = 1
   '
   Call FINAL("*LIDAPIBC")
   '
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    '
    Call ABRECABAN
    
    'CueCorri.Execute "update cajabanc set mone_emis=1 where CODICOM_LAR = 'NC-0004-00000570-A' and nuclien=6"
    'CueCorri.Execute "update cajabanc set mone_emis=1 where CODICOM_LAR = 'ND-0004-00000056-A' and nuclien=6"
    CueCorri.Execute "update cajabanc set impo_dola=0 where CODICOM_LAR ='NC-0001-00000134-A'"
    CueCorri.Execute "update cajabanc set fehoreal='14/06/13 17:58:00' where CODICOM_LAR = 'FC-0004-00000828-A' and nuclien=6"
    CueCorri.Execute "update cajabanc set mone_emis=1 where CODICOM_LAR ='NC-0001-00000134-A'"
    CueCorri.Execute "update cajabanc set impo_dola=0 where CODICOM_LAR ='NC-0001-00000134-A'"
    CueCorri.Execute "update cajabanc set impo_dola=0 where CODICOM_LAR ='NC-0004-00000570-A'"
    CueCorri.Execute "update cajabanc set mone_emis=1 where CODICOM_LAR ='NC-0004-00000570-A'"
    CueCorri.Execute "update cajabanc set impo_dola=0 where CODICOM_LAR ='ND-0004-00000056-A'"
    CueCorri.Execute "update cajabanc set mone_emis=1 where CODICOM_LAR ='ND-0004-00000056-A'"
    CueCorri.Execute "update cajabanc set impo_dola=0 where CODICOM_LAR ='FC-0001-00004954-A'"
    CueCorri.Execute "update cajabanc set mone_emis=1 where CODICOM_LAR ='FC-0001-00004954-A'"
'    With SaDeuPen
'      .FindFirst "CODICOM_LAR = 'NC-0004-00000570-A' and nuclien=6"
'10    If .NoMatch = False Then
'        .Edit
'        A = !nuclien
'        !SaldDebe_Dola = 0
'        !SaldAcre_Dola = 0
'        !Mone_Emis = 1
'        .Update
'        .MoveNext
'        .FindNext "CODICOM_LAR = 'NC-0004-00000570-A' and nuclien=6"
'        GoTo 10
'      End If
'    End With
'
'    With SaDeuPen
'      .FindFirst "CODICOM_LAR = 'ND-0004-00000056-A'"
'20    If .NoMatch = False Then
'        .Edit
'        !SaldDebe_Dola = 0
'        !SaldAcre_Dola = 0
'        !Mone_Emis = 1
'        .Update
'        .FindNext "CODICOM_LAR = 'ND-0004-00000056-A'"
'        GoTo 20
'      End If
'    End With
'
'    With SaDeuPen
'      .FindFirst "CODICOM_LAR = 'NC-0001-00000134-A'"
'30    If .NoMatch = False Then
'        .Edit
'        !SaldDebe_Dola = 0
'        !SaldAcre_Dola = 0
'        !Mone_Emis = 0
'        .Update
'        .FindNext "CODICOM_LAR = 'NC-0001-00000134-A'"
'        GoTo 3
'      End If
'    End With

    FIN& = ULTREG&("CARBIMON")
    For KUL& = 1 To FIN&
      XXC$ = REGLEIDO$("CARBIMON", KUL&)
      COCOLAR$ = TRIM$(Mid$(XXC$, 13, 18))
      If COCOLAR$ = "ND-0004-00000056-A" Or COCOLAR$ = "FC-0001-00004954-A" Or InStr(1, COCOLAR$, "0000134") > 0 Then
            Call REPLA(XXC$, Space(18), 13, 18)
            Call GRAREG("CARBIMON", XXC$, KUL&)
      End If
      COCOLAR$ = TRIM$(Mid$(XXC$, 45, 18))
      If COCOLAR$ = "ND-0004-00000056-A" Or COCOLAR$ = "FC-0001-00004954-A" Or InStr(1, COCOLAR$, "0000134") > 0 Then
            Call REPLA(XXC$, Space(18), 45, 18)
            Call GRAREG("CARBIMON", XXC$, KUL&)
      End If
    Next KUL&

    Call RECUEC04.REAPLICRE(6)

End Sub

'Private Sub Command6_Click()

'PONE POR DEFECTO EN TABLA DE CONFIGURACION DE INGRESOSO BRUTOS A TODOS EN POSICION 1 PORCENTAJE 100 Y VACIO AL MOTIVO DE EXENTO
'   Command6.Enabled = False
'
'   FIN& = ULTREG&("DARIBPRO")
'   For I& = 1 To FIN&
'        X$ = REGLEIDO$("DARIBPRO", I&)
'        Call REPLA(X$, Chr$(1), 3, 1)
'        Call REPLA(X$, MKS(100), 5, 4)
'        Call REPLA(X$, Space$(32), 9, 32)
'        Call GRAREG("DARIBPRO", X$, I&)
'   Next I&
'
'   FIN& = ULTREG&("DAPIBCLI")
'   For I& = 1 To FIN&
'        X$ = REGLEIDO$("DAPIBCLI", I&)
'        Call REPLA(X$, Chr$(1), 3, 1)
'        Call REPLA(X$, MKS(100), 5, 4)
'        Call REPLA(X$, Space$(32), 9, 32)
'        Call GRAREG("DAPIBCLI", X$, I&)
'   Next I&
'
'   Command6.Enabled = True
'
'End Sub

'
Private Sub Form_Load()
  '
  If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Image2.Visible = False
      OPMASCL04.Frame6.BackColor = &H8000000F
      OPMASCL04.Frame1.BackColor = &H8000000F
      OPMASCL04.Frame10.BackColor = &H8000000F
      OPMASCL04.Frame11.BackColor = &H8000000F
      OPMASCL04.Frame12.BackColor = &H8000000F
      OPMASCL04.Frame2.BackColor = &H8000000F
      OPMASCL04.Frame3.BackColor = &H8000000F
      OPMASCL04.Frame4.BackColor = &H8000000F
      
      OPMASCL04.Frame7.BackColor = &H8000000F
      OPMASCL04.Frame5.BackColor = &H8000000F
      OPMASCL04.Frame8.BackColor = &H8000000F
      OPMASCL04.Frame9.BackColor = &H8000000F
      For I% = 4 To OPMASCL04.Option1.UBound
         OPMASCL04.Option1(I%).BackColor = &H8000000F
      Next I%
      OPMASCL04.Option2.BackColor = &H8000000F
      OPMASCL04.Option3.BackColor = &H8000000F
      OPMASCL04.Check1.BackColor = &H8000000F
      Me.BackColor = &H8000000F
  End If

  Screen.MousePointer = 11
  Call CARTAVALI
  Screen.MousePointer = 1
  '
  TABLA$(0) = "TATIMAT"       ' tabla de tipos de material
  TABLA$(1) = "*UNIMED"
  TABLA$(2) = "TADEPOSI"
  TABLA$(3) = ""       ' grupos contables por articulo
  TABLA$(4) = "*VENDEDOR"
  TABLA$(5) = "TAPROVIN"
  TABLA$(6) = "CONPAGO"
  TABLA$(7) = "LISPRE04"
  TABLA$(8) = "CUEINGAS"
  TABLA$(9) = "GRUCOCLI"
  TABLA$(10) = "ZONAREP"
  TABLA$(11) = "GRUFACLI"
  '
  PRONUCLI$ = CONTROL$("", "PRONUCLI")
  If XVALO(PRONUCLI$) >= 0 Then
    Option3.Value = True
    Text3 = PRONUCLI$
  End If
  '
  For Index = 0 To 7
    VAPRE$ = TRIM$(CONTROL$("AMACLI", ARCHE$(Index)))
    COTEXT(Index) = "": DETEXT(Index) = ""
    If VAPRE$ <> "" Then
      TV11% = TIPFIELD(ARCHE$(Index), 1)
      LF11% = LENFIELD(ARCHE$(Index), 1)
      ARGU$ = CBIN$(VAPRE$, TV11%, LF11%)
      ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
      If ECO$ <> "" Then
          COTEXT(Index).TEXT = VAPRE$
          DETEXT(Index).TEXT = ECO$
        Else
          COTEXT(Index).TEXT = ""
          DETEXT(Index).TEXT = ""
       End If
    End If
    DETEXT(Index).Refresh
    DETEXT(Index).TabStop = False
  Next Index
  ' '\\\OT-05-0198-WAR-19/04/05-///
  If CONTROL$("FORMCUIT", "SUPRIVAL") = "SI" Then
      Check1.Value = 1
  End If
  '''\\\OT-05-0198-WAR-FIN///
End Sub
'
Private Sub Option1_Click(Index As Integer)
    Option1(Index).Value = True
    Text2.TEXT = Option1(Index).Caption
    MODUCONE$ = TABLA$(Index)
    TETABLA$ = MODUCONE$
    If Left$(TETABLA$, 1) = "#" Then
        TETABLA$ = Mid$(TETABLA$, 2) + "(#)"
      ElseIf Left$(TETABLA$, 1) = "*" Then
        TETABLA$ = Mid$(TETABLA$, 2) + "(*)"
    End If
    Text1.TEXT = TETABLA$
    '
    For K% = 4 To 10
       If K% <> Index Then
          Option1(K%).Value = False
       End If
    Next K%
    'Command16.SetFocus
End Sub

Private Sub Option1_DblClick(Index As Integer)
    If TRIM$(MODUCONE$) <> "" Then
        If TRIM$(MODUCONE$) = "GRUCOCLI" Then
            LDT$ = LUDAT$
            Call CIERRARCH("GRUCOCLI")
            Call CIERRARCH("!GRUCOCLI")
            'Call COPYFILE(LDT$ + "GRUCOCLI.DAT", "C:\FLEXOFT\GRUCOCLI.WKF")
            Call TRALOCAL("GRUCOCLI", "")
            VTB% = VENTABU%("GRUCOCLI")
            If VTB% < 0 Then Exit Sub
            J& = 0
            For I& = 1 To ULTREG&("!GRUCOCLI")
                X$ = REGLEIDO$("!GRUCOCLI", I&)
                If CVI(Left$(X$, 2)) > 0 Then
                    J& = J& + 1
                    Call GRAREG("GRUCOCLI", X$, J&)
                End If
            Next I&
            Call SETULTREG("GRUCOCLI", J&)
            Call CIERRARCH("GRUCOCLI")
          ElseIf MODUCONE = "TADEPOSI" Then
            Call CARTADEPOSI
          ElseIf MODUCONE = "TATIMAT" Then
            Call CARTATIMA
          ElseIf MODUCONE = "CONPAGO" Then
            Call CARCONPAGOS
          ElseIf MODUCONE = "TAPROVIN" Then
            Call CARTAPROVIN
          ElseIf MODUCONE = "CUEINGAS" Then
            Call CARCUEINGAS
          ElseIf MODUCONE = "ZONAREP" Then
            Call CARZONAREP
          ElseIf MODUCONE = "GRUFACLI" Then
            DEFDESCO.Show 1
          ElseIf MODUCONE = "*VENDEDOR" Then
                 RTA% = ALTERNA%("ABM Vendedores\Datos Complementarios")
              If RTA% = 1 Then
                  Call CONECRUN(MODUCONE$, Text2.TEXT)
                  Call FILESORT("VENDEDOR", "", 1, 1, "ASC")
                  ElseIf RTA% = 2 Then
25                        Call SELECHO("VENDEDOR")
                          OP$ = VALACT1$("VENDEDOR")
                          If OP$ <> "" Then
                                 X$ = "": Y$ = ""
                                 NUMVEN% = XVALO(OP$)
                                 NOVEN$ = VALACT2$("VENDEDOR")
                                 X$ = FILTERGETRECORD("DIMAIVEN", 1, Chr$(NUMVEN%))
                                 VDEF$ = ""
                                 Call REPLA(VDEF$, Mid$(X$, 65, 48), 1, 48)
                                 Call REPLA(VDEF$, Mid$(X$, 113, 48), 49, 48)
                                 Call REPLA(VDEF$, Mid$(X$, 161, 32), 97, 32)
                                 Call REPLA(VDEF$, Mid$(X$, 193, 48), 129, 32)
                                 Call REPLA(VDEF$, Mid$(X$, 225, 48), 161, 48)
                                 Call REPLA(VDEF$, Mid$(X$, 2, 63), 209, 63)
                                 
                                 VDEF$ = OBJPLA$("DATCOVEND", VDEF$)
                                 If VDEF$ <> "" Then
                                    Call REPLA(X$, Chr$(NUMVEN%), 1, 1)
                                    Call REPLA(X$, Mid$(VDEF$, 209, 63), 2, 63)
                                    Call REPLA(X$, Mid$(VDEF$, 1, 48), 65, 48)
                                    Call REPLA(X$, Mid$(VDEF$, 49, 48), 113, 48)
                                    Call REPLA(X$, Mid$(VDEF$, 97, 32), 161, 32)
                                    Call REPLA(X$, Mid$(VDEF$, 129, 32), 193, 32)
                                    Call REPLA(X$, Mid$(VDEF$, 161, 48), 225, 48)
                                    Call FILTERSETRECORD("DIMAIVEN", 1, Chr$(NUMVEN%), X$)
                                 End If
                                 GoTo 25
                              Else
                                 Exit Sub
                          End If
                  Else: Exit Sub
              End If
        Else
            Call CONECRUN(MODUCONE$, Text2.TEXT)
        End If
    End If

End Sub


Private Sub Option2_Click()
   Option2.Value = True
   Text3.TEXT = ""
End Sub

Private Sub Option3_Click()
   Option3.Value = True
End Sub

Private Sub Text1_GotFocus()
   Command16.SetFocus
End Sub

Private Sub Text2_gotfocus()
   Command16.SetFocus
End Sub

Private Sub Text3_Click()
   Option3.Value = True
End Sub



Private Sub COTEXT_DblClick(Index As Integer)
    Call SELECHO(ARCHE$(Index))
    If TRIM$(VALACT1$(ARCHE$(Index))) <> "" Then
        COTEXT(Index).TEXT = VALACT1$(ARCHE$(Index))
        COTEXT(Index).Refresh
        DETEXT(Index).TEXT = VALACT2$(ARCHE$(Index))
        DETEXT(Index).Refresh
        DETEXT(Index).TabStop = False
    End If
    COTEXT(Index).SetFocus
End Sub
'
Private Sub BOTSEL_Click(Index As Integer)
    Call SELECHO(ARCHE$(Index))
    If TRIM$(VALACT1$(ARCHE$(Index))) <> "" Then
        COTEXT(Index).TEXT = VALACT1$(ARCHE$(Index))
        COTEXT(Index).Refresh
        DETEXT(Index).TEXT = VALACT2$(ARCHE$(Index))
        DETEXT(Index).Refresh
        DETEXT(Index).TabStop = False
    End If
    COTEXT(Index).SetFocus
End Sub
'
Private Sub COTEXT_LOSTFOCUS(Index As Integer)
    If TRIM$(COTEXT(Index).TEXT) <> "" Then
        TV11% = TIPFIELD(ARCHE$(Index), 1)
        LF11% = LENFIELD(ARCHE$(Index), 1)
        ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
        ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
        If ECO$ <> "" Then
             COTEXT(Index).TEXT = TRIM$(COTEXT(Index).TEXT)
             DETEXT(Index).TEXT = ECO$
             DETEXT(Index).Refresh
             DETEXT(Index).TabStop = False
           Else
             DETEXT(Index).TEXT = ""
             Exit Sub
        End If
      Else
        DETEXT(Index).TEXT = ""
    End If
End Sub
'

Private Sub COTEXT_KeyPress(Index As Integer, KeyAscii As Integer)
    '
    If KeyAscii = 13 Then
       If TRIM$(COTEXT(Index).TEXT) <> "" Then
          TV11% = TIPFIELD(ARCHE$(Index), 1)
          LF11% = LENFIELD(ARCHE$(Index), 1)
          ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
          ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
          If ECO$ <> "" Then
               DETEXT(Index).TEXT = ECO$
               DETEXT(Index).Refresh
               DETEXT(Index).TabStop = False
             Else
               DETEXT(Index).TEXT = ""
               COTEXT(Index).SetFocus
          End If
       End If
    End If
End Sub

Private Sub DETEXT_CLICK(Index As Integer)
    COTEXT(Index).SetFocus
End Sub
Sub CARZONAREP()
    If TRIM$(MODUCONE$) = "ZONAREP" Then
      LDT$ = LUDAT$
      Call CIERRARCH("ZONAREP")
      Call CIERRARCH("!ZONAREP")
      Call TRALOCAL("ZONAREP", "")
      VTB% = VENTABU%("ZONAREP")
      If VTB% < 0 Then GoTo 99
      J& = 0
      For I& = 1 To ULTREG&("!ZONAREP")
          X$ = REGLEIDO$("!ZONAREP", I&)
          If TRIM$(Left$(X$, 8)) <> "" Then
              J& = J& + 1
              Call GRAREG("ZONAREP", X$, J&)
          End If
      Next I&
      Call SETULTREG("ZONAREP", J&)
      Call CIERRARCH("ZONAREP")
    End If
99 End Sub

            

