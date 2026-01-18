VERSION 5.00
Begin VB.Form STPCOS07 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Costos - Formularios"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   11460
   FillColor       =   &H00808080&
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
   ScaleHeight     =   5190
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Categorías de Clientes por  Mano de Obra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   5520
      TabIndex        =   40
      Top             =   120
      Width           =   4920
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1965
         Left            =   240
         TabIndex        =   46
         Top             =   1560
         Width           =   4425
         Begin VB.CommandButton Command13 
            Caption         =   "Acceso Detalle de Gastos"
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
            Left            =   240
            Picture         =   "STPCOS07.frx":0000
            TabIndex        =   52
            ToolTipText     =   "Actualizacion de Precios de Lista por Costo Calculado"
            Top             =   1200
            Width           =   2685
         End
         Begin VB.CommandButton Command12 
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
            Left            =   3080
            Picture         =   "STPCOS07.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Grabar Categoría"
            Top             =   1200
            Width           =   1245
         End
         Begin VB.TextBox Text6 
            BackColor       =   &H8000000F&
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
            Left            =   2775
            Locked          =   -1  'True
            TabIndex        =   48
            Text            =   " "
            Top             =   360
            Width           =   1380
         End
         Begin VB.CommandButton Command11 
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
            Left            =   4140
            TabIndex        =   47
            Top             =   360
            Width           =   175
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Categoría:"
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
            Left            =   360
            TabIndex        =   50
            Top             =   420
            Width           =   2325
         End
         Begin VB.Label Label122 
            Alignment       =   1  'Right Justify
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
            Left            =   1800
            TabIndex        =   49
            Top             =   825
            Width           =   2535
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   240
         TabIndex        =   41
         Top             =   240
         Width           =   4440
         Begin VB.CommandButton Command10 
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
            Left            =   4140
            TabIndex        =   43
            Top             =   360
            Width           =   175
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
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
            Left            =   2775
            TabIndex        =   42
            Text            =   " "
            Top             =   360
            Width           =   1380
         End
         Begin VB.Label Label11 
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
            Left            =   120
            TabIndex        =   45
            Top             =   825
            Width           =   4215
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cta.Cliente"
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
            Left            =   360
            TabIndex        =   44
            Top             =   420
            Width           =   2325
         End
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10560
      Top             =   2520
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Formulario de Impresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   210
      TabIndex        =   32
      Top             =   1365
      Width           =   5160
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
         Height          =   730
         Left            =   4095
         Picture         =   "STPCOS07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   300
         Width           =   855
      End
      Begin VB.CommandButton Command7 
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
         Left            =   3780
         TabIndex        =   34
         Top             =   315
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
         Height          =   300
         Left            =   2415
         TabIndex        =   33
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label8 
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
         TabIndex        =   37
         Top             =   375
         Width           =   2325
      End
      Begin VB.Label Label7 
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
         Left            =   210
         TabIndex        =   36
         Top             =   735
         Width           =   3750
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Parámetros de Cálculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   20
      Top             =   2760
      Width           =   5160
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
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
         Left            =   3795
         TabIndex        =   53
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text41 
         Alignment       =   1  'Right Justify
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
         Left            =   3795
         TabIndex        =   27
         Top             =   1140
         Width           =   855
      End
      Begin VB.TextBox Text39 
         Alignment       =   1  'Right Justify
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
         Left            =   3795
         TabIndex        =   26
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text43 
         Alignment       =   1  'Right Justify
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
         Left            =   3795
         TabIndex        =   25
         Top             =   1740
         Width           =   855
      End
      Begin VB.TextBox Text36 
         Alignment       =   1  'Right Justify
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
         Left            =   3795
         TabIndex        =   22
         Top             =   825
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3435
         TabIndex        =   21
         Top             =   420
         Width           =   1200
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Precio Minorista (%):"
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
         Index           =   0
         Left            =   1560
         TabIndex        =   54
         Top             =   2085
         Width           =   2115
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Scrap (%):"
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
         Index           =   22
         Left            =   1260
         TabIndex        =   30
         Top             =   1185
         Width           =   2430
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Utilidad (%):"
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
         Index           =   29
         Left            =   1785
         TabIndex        =   29
         Top             =   1485
         Width           =   1905
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Impuestos (%):"
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
         Index           =   30
         Left            =   1575
         TabIndex        =   28
         Top             =   1785
         Width           =   2115
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gastos Generales (%):"
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
         Index           =   2
         Left            =   105
         TabIndex        =   24
         Top             =   825
         Width           =   3585
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Hora  Mano de Obra ($):"
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
         Left            =   480
         TabIndex        =   23
         Top             =   465
         Width           =   2955
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Listas de Precios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   5520
      TabIndex        =   10
      Top             =   3840
      Width           =   4920
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
         Height          =   300
         Left            =   1035
         TabIndex        =   18
         Text            =   " "
         Top             =   420
         Width           =   435
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
         Height          =   300
         Left            =   1035
         TabIndex        =   17
         Text            =   " "
         Top             =   840
         Width           =   435
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
         Height          =   300
         Left            =   1455
         TabIndex        =   13
         Top             =   840
         Width           =   175
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
         Height          =   300
         Left            =   1455
         TabIndex        =   11
         Top             =   420
         Width           =   175
      End
      Begin VB.Label Label6 
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
         Left            =   1770
         TabIndex        =   19
         Top             =   840
         Width           =   3075
      End
      Begin VB.Label Label5 
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
         Left            =   1770
         TabIndex        =   16
         Top             =   420
         Width           =   3075
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dólares:"
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
         Left            =   -1380
         TabIndex        =   14
         Top             =   900
         Width           =   2325
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pesos:"
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
         Left            =   -1380
         TabIndex        =   12
         Top             =   480
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004040&
      Height          =   8520
      Left            =   10560
      ScaleHeight     =   8460
      ScaleWidth      =   1965
      TabIndex        =   5
      Top             =   120
      Width           =   2025
      Begin VB.CommandButton Command21 
         Caption         =   "Tbls"
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
         Left            =   120
         Picture         =   "STPCOS07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Tabla de Sectores"
         Top             =   3480
         Width           =   650
      End
      Begin VB.CommandButton Command14 
         Caption         =   "$$"
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
         Picture         =   "STPCOS07.frx":08A8
         TabIndex        =   38
         ToolTipText     =   "Actualizacion de Precios de Lista por Costo Calculado"
         Top             =   1860
         Width           =   650
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
         Height          =   640
         Left            =   105
         Picture         =   "STPCOS07.frx":09F2
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Cotización Dolar y Tipo de Cambio"
         Top             =   2700
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   100
         MouseIcon       =   "STPCOS07.frx":0CFC
         Picture         =   "STPCOS07.frx":1006
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1050
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
         Height          =   640
         Left            =   100
         Picture         =   "STPCOS07.frx":1310
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   300
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   100
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   6
         Top             =   4320
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   7
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "STPCOS07.frx":145A
         Top             =   4560
         Width           =   1515
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Formulario de Impresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   210
      TabIndex        =   0
      Top             =   105
      Width           =   5160
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
         Height          =   300
         Left            =   2415
         TabIndex        =   4
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
         Height          =   300
         Left            =   3780
         TabIndex        =   3
         Top             =   315
         Width           =   175
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
         Height          =   730
         Left            =   4095
         Picture         =   "STPCOS07.frx":337C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label4 
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
         Left            =   210
         TabIndex        =   15
         Top             =   735
         Width           =   3750
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
         Top             =   375
         Width           =   2325
      End
   End
   Begin VB.Menu mnuTransacciones 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuprobra 
         Caption         =   "Parámetros de Proyección de Mano de Obra"
      End
   End
End
Attribute VB_Name = "STPCOS07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LPREX%
Public CONDI_PUBLIC$
Private Sub Command1_Click()
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   Call SELECHO("DEUDOR1")
   Text3 = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call SELECHO("GASTOGEN/Gastos Generales")
   Text6 = VALACT1$("GASTOGEN")
'   Label122 = VALACT2$("GASTOGEN")
   Command11.Enabled = True

End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   If Label11 = "" Then
      Call COMUNI("Falta Ingresar Cliente o Cta. No Válida")
      GoTo 99
   End If
   If Label122 = "" Then
      Call COMUNI("Falta Ingresar Categoría o es Inexistente")
      GoTo 99
   End If
   '
   CONDI$ = "Nume_Cli = " & XVALO(Text3)
   Call ACTUREGISTRO("Deudor12", "CateMobra", CONDI$, TRIM$(Text6), REGAF&, "NOMESS")
   '
   Call COMUNI("Categoría Registrada")
   '
99 Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   'TEXTO CTAYCOST
   'boton de setup de anacos frm stpcos07
    Command13.Enabled = False
    '
    CANTICUENTAS% = XVALO(CantRegistros("PLANCUEN", , "NOMESS"))
    Dim FILAXX&()
    ReDim Preserve FILAXX&(CANTICUENTAS%, 3)
    If EXISTE%(LUDAT$ & "TACENCOS") < 1 Then GoTo 99
    FIN& = ULTREG&("TACENCOS")
    If FIN& < 1 Then GoTo 99
    
    MUESTRADATOS1.Height = 3840
    MUESTRADATOS1.Command1.Visible = True
10  Call CARGA_ENCABEZADO(MUESTRADATOS1.GRID, "Centro de Costos/f18;Denomincación/A24;Importe/F20", MUESTRADATOS1)
    
    ANOX_ACTUAL% = Year(Date)
    MES% = 0
    ANOSELECCIONADO% = ANOX_ACTUAL
15  OPX% = COMALTER%("Modo de Selección\\Elegir Año\Seleccionar Mes y Año")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    If OPX% = 1 Then 'ELIGE POR ANO
        ANOSELECCIONADO% = XVALO(InputBox$("Selección de Año de Informe", "Ingrese Año (yyyy)", ANOSELECCIONADO%))
        If ANOSELECCIONADO% < 1980 Or ANOSELECCIONADO% > ANOX_ACTUAL% Then GoTo 99
        ANO2% = Right$(TRIM$(ANOSELECCIONADO%), 2)
    ElseIf OPX% = 2 Then 'ELIGE MES Y ANO
        Call SELMESA(Des%, Has%, PERIO$)
        If Des% > 30 Then
            Desde = FETEXCOR1$(Des%)
            ANOSELECCIONADO% = Year(Desde)
            MES% = Month(Desde)
            Desde = FETEXCOR1$(Des%)
            Hasta = FETEXCOR1$(Has%)
        Else
            GoTo 15
        End If
    End If
    
    '
    For k& = 1 To FIN&
       XX$ = REGLEIDO$("TACENCOS", k&)
       CODCCOSTO$ = TRIM$(Mid$(XX$, 1, 8))
       DENOMINA$ = TRIM$(Mid$(XX$, 9, 40))
       MUESTRADATOS1.GRID.Rows = MUESTRADATOS1.GRID.Rows + 1
       MUESTRADATOS1.GRID.TextMatrix(k&, 1) = CODCCOSTO$
       MUESTRADATOS1.GRID.TextMatrix(k&, 2) = DENOMINA$
       MUESTRADATOS1.GRID.TextMatrix(k&, 3) = FORMATNUM$(IMPO_ACUMU_CENCOS#(XVALO(CODCCOSTO$), ANOSELECCIONADO%, MES%), "F14.2")
    Next k&
    '
    MUESTRADATOS1.Caption = MUESTRADATOS1.Caption & SAFETEXT$(ANOSELECCIONADO%)
    Call COLOREAR_GRILLASQL(MUESTRADATOS1.GRID, &HE0E0E0)
    
20  MUESTRADATOS1.Show 1

    '
    If MUESTRADATOS1.RESP_ZELE$ = "SELECCIONMULTIPLE" Then Call SeleccionMultipleCentroCostos: GoTo 20
    CCOSSTRING$ = MUESTRADATOS1.RESP_ZELE$
    
    CCOSTX% = XVALO(CCOSSTRING$)
    If CCOSTX% < 1 Then GoTo 10
    DENOMICENCOS$ = TRIM$(ECOARCH("TACENCOS", TRIM$(CCOSSTRING$)))
    
    Call CARGA_ENCABEZADO(MUESTRADATOS2.GRID, "Cta.Cble./A13;Denomincación/A24;Enero " & ANO2% & "/F12.2;Febrero " & ANO2% & "/F12.2;Marzo " & ANO2% & "/F12.2;Abril " & ANO2% & "/F12.2;Mayo " & ANO2% & "/F12.2;Junio " & ANO2% & "/F12.2;Julio " & ANO2% & "/F12.2;Agosto " & ANO2% & "/F12.2;Septiembre " & ANO2% & "/F12.2;Octurbre " & ANO2% & "/F12.2;Noviembre " & ANO2% & "/F12.2;Diciembre " & ANO2% & "/F12.2;Total/F14.2;CODIEMPR/F0", MUESTRADATOS2, 1, 1)
    '
    'CARGO ENCABEZADO DE GRILLA DE FILTRO
    MUESTRADATOS2.MSF_2.Rows = 3
    MUESTRADATOS2.MSF_2.FixedRows = 2
    Call CARGA_ENCABEZADO(MUESTRADATOS2.MSF_2, "Cta.Cble./A13;Denomincación/A24;Enero " & ANO2% & "/F12.2;Febrero " & ANO2% & "/F12.2;Marzo " & ANO2% & "/F12.2;Abril " & ANO2% & "/F12.2;Mayo " & ANO2% & "/F12.2;Junio " & ANO2% & "/F12.2;Julio " & ANO2% & "/F12.2;Agosto " & ANO2% & "/F12.2;Septiembre " & ANO2% & "/F12.2;Octurbre " & ANO2% & "/F12.2;Noviembre " & ANO2% & "/F12.2;Diciembre " & ANO2% & "/F12.2;Total/F14.2;CODIEMPR/F0", MUESTRADATOS2, 1, 1)
    '
    Dim Total(13)
    Erase Total
    MUESTRADATOS2.GRID.Rows = 2
    IMPOTOTAL# = 0
    VUELTA% = 0
    SQLX$ = "SELECT P.CODEXCUE ,  P.Denocue, C.CenCosto, SUM(C.IDebeComp) AS FACSUMAIDEBE, SUM(C.IHabeComp) AS NCSUMAHABER,"
    SQLX$ = SQLX$ + " MONTH(C.FECHEMISI) as mes, YEAR(C.FECHEMISI) AS ANO ,C.CodiEmpr,C.cuecontai"
    SQLX$ = SQLX$ + " FROM CAJABANC C  WITH(NOLOCK) INNER JOIN PLANCUEN P WITH(NOLOCK)"
    SQLX$ = SQLX$ + " ON C.cuecontai = P.CodInCue"
    'CONDICION
    SQLX$ = SQLX$ + " Where c.NuProve > 0 And c.CENCOSTO = " & CCOSTX% & " AND YEAR(C.FECHEMISI) = " & ANOSELECCIONADO%
    '
    If MES% > 0 Then SQLX$ = SQLX$ + " AND MONTH(FECHEMISI)= " & MES%
    SQLX$ = SQLX$ + " GROUP BY  C.CenCosto,YEAR(C.FECHEMISI), MONTH(C.FECHEMISI),C.cuecontai, P.CODEXCUE , P.Denocue,C.CodiEmpr"
    SQLX$ = SQLX$ + " ORDER BY YEAR(C.FECHEMISI), MONTH(C.FECHEMISI),C.cuecontai, P.CODEXCUE, P.Denocue, C.CodiEmpr  "
    'Set RS = READSET(SQLX$)
    Set RS = FLEXCONN.Execute(SQLX$)

    
    With RS
    i% = 1
    IMPOTOTAL# = 0
    Do While Not .EOF
      MEX = XVALO(!MES)
      SUBRAYAR% = 0
      If VUELTA% = 0 Then FILAXX&(XVALO(!CUECONTAI), XVALO(!CODIEMPR)) = i%
      If ((SAFETEXT$(!CodExCue) <> CODEXCUE_ANT) Or (XVALO(!CODIEMPR) <> CodiEmpr_ANT)) And (VUELTA% > 0) Then
           i% = i% + 1
           MUESTRADATOS2.GRID.Rows = i%
           MUESTRADATOS2.GRID.Rows = MUESTRADATOS2.GRID.Rows + 1
           IMPOTOTAL# = 0
           'USO UN VECTOR PARA LLEVAR EL CONTROL DE LA FILA YA QUE CUANDO CAMBIABA DE MES LA MISMA CUENTA LO IMPRIMIA EN OTRA FILA
           If FILAXX&(XVALO(!CUECONTAI), XVALO(!CODIEMPR)) = 0 Then
             FILAXX&(XVALO(!CUECONTAI), XVALO(!CODIEMPR)) = i%
           End If
           
      End If
      '
      If XVALO(!CODIEMPR) > 0 Then SUBRAYAR% = 1 ' PARA SUBRAYAR CASO MULTIEMPRESA  A MODO DE DIFERENCIAR UNO Y OTRO
      A = FILAXX&(XVALO(!CUECONTAI), 1)
      B = XVALO(!CODIEMPR)
      MUESTRADATOS2.GRID.Row = FILAXX&(XVALO(!CUECONTAI), XVALO(!CODIEMPR)): MUESTRADATOS2.GRID.COL = 1
      MUESTRADATOS2.GRID.TextMatrix(FILAXX&(XVALO(!CUECONTAI), XVALO(!CODIEMPR)), 1) = SAFETEXT$(!CodExCue): If SUBRAYAR% > 0 Then MUESTRADATOS2.GRID.CellFontUnderline = True: MUESTRADATOS2.GRID.CellForeColor = vbBlue
      MUESTRADATOS2.GRID.TextMatrix(FILAXX&(XVALO(!CUECONTAI), XVALO(!CODIEMPR)), 2) = SAFETEXT$(!DENOCUE)
        IMPO# = XVALO(!FACSUMAIDEBE) - XVALO(!NCSUMAHABER)
      MUESTRADATOS2.GRID.TextMatrix(FILAXX&(XVALO(!CUECONTAI), XVALO(!CODIEMPR)), MEX + 2) = TRIM$(FORMATNUM$(IMPO#, "F12.2"))
        IMPOTOTAL# = IMPOTOTAL# + IMPO#
      'MUESTRADATOS2.GRID.TextMatrix(FILAXX&(XVALO(!CUECONTAI), XVALO(!CodiEmpr)), 15) = FORMATNUM$(IMPOTOTAL#, "F12.2") ' LO SUMA AL FINAL
      Total(MEX) = Total(MEX) + IMPO#
      'Total(13) = Total(13) + IMPOTOTAL# 'LO SUMA AL FINAL
      MUESTRADATOS2.GRID.TextMatrix(FILAXX&(XVALO(!CUECONTAI), XVALO(!CODIEMPR)), 16) = XVALO(!CODIEMPR)
      '
      'CUANDO CAMBIA LA CUENTA CONTABLE PASA AL SIGUIENTE REGISTRO
      CODEXCUE_ANT = TRIM$(!CodExCue)
      CodiEmpr_ANT = XVALO(!CODIEMPR)
      VUELTA% = 1
      .MoveNext
    Loop
    End With

    'Call COLOREAR_GRILLASQL_VERTICAL(MUESTRADATOS2.GRID, &HE0E0E0, 0, 2)
    MUESTRADATOS2.Caption = "Informe de Gastos - Centro de Costos: " & DENOMICENCOS$ & " (Año: " & ANOSELECCIONADO% & ")"
    Screen.MousePointer = 1
    MUESTRADATOS2.GRID.FixedCols = 3 '
    '
    'ORDENO POR LA DENOMINACION DE LA CUENTA.
    MUESTRADATOS2.GRID.COL = 1
    MUESTRADATOS2.GRID.Sort = 5
    
    'IMPRIMIR TOTALES VERTICALES

    i% = MUESTRADATOS2.GRID.Rows
    'TOTALIZO AL FINAL POR COLUMNA
    MUESTRADATOS2.GRID.Rows = i% + 2
     MUESTRADATOS2.GRID.TextMatrix(i%, 2) = "TOTALES: "
    For J% = 3 To 15 ' 3 A 15 POR QUE LOS MESES ARRANCAN DE LA COLUMNA 3
       MUESTRADATOS2.GRID.TextMatrix(i%, J%) = FORMATNUM$(Total(J% - 2), "F12.2") ' MENOS 2 POR EL QUE EL VECTOR TIENE EL NUMERO REAL DEL MES
    Next J%
    '
    'TOTALIZA DE MANERA HORIZONTAL
    For L% = 1 To MUESTRADATOS2.GRID.Rows - 1
        TOTALFILA# = 0
        For R% = 3 To 15
          TOTALFILA# = TOTALFILA# + XVALO(MUESTRADATOS2.GRID.TextMatrix(L%, R%))
        Next R%
        MUESTRADATOS2.GRID.TextMatrix(L%, 15) = FORMATNUM$(TOTALFILA#, "F12.2")
    Next L%
    
    Call COLOREAR_GRILLASQL(MUESTRADATOS2.GRID, &HE0E0E0)

    'COLOREO LA FILA DE LOS TOTALES
    Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MUESTRADATOS2.GRID, &HC00000, i%)
    Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MUESTRADATOS2.GRID, &H808080, i%)
    '
    'POSICIONARSE AL FINAL DE LA LISTA
     TPXI& = MUESTRADATOS2.GRID.Rows - 22: If TPXI& < 1 Then TPXI& = 1
     On Error Resume Next
     MUESTRADATOS2.GRID.TopRow = TPXI&
     On Error GoTo 0
     '
    MUESTRADATOS2.ANO = SAFETEXT$(ANOSELECCIONADO%)
    MUESTRADATOS2.CENTROCOSTO = CCOSSTRING$
    MUESTRADATOS2.Show 1
    'LIMPIO EL VECTOR
    For ii% = 0 To UBound(FILAXX&)
      For J% = 0 To 3
       FILAXX&(ii%, J%) = 0
      Next J%
    Next ii%
    
    GoTo 20
    '
    'FALTA VER EL TEMA DE LOS TOTALES
99
   Command13.Enabled = True

End Sub
Function IMPO_ACUMU_CENCOS#(CCSTO%, ANOX%, MES%)

        SQLX$ = " SELECT SUM(IHabeComp) AS HABE, SUM (IDebeComp) DEBE, CENCOSTO FROM CAJABANC"
        SQLX$ = SQLX$ + " Where CENCOSTO = " & CCSTO%
        SQLX$ = SQLX$ + " And Year(FECHEMISI) = " & ANOX%
        If MES% > 0 Then SQLX$ = SQLX$ + " And MONTH(FECHEMISI) = " & MES%
        SQLX$ = SQLX$ + "  GROUP BY CENCOSTO"
        Dim RS As New ADODB.Recordset
        Set RS = FLEXCONN.Execute(SQLX$)
        IMPO_ACUMU_CENCOS# = XVALO(RS!DEBE) - XVALO(RS!HABE)
        RS.Close
        Set RS = Nothing
        
End Function

Private Sub Command14_Click()
   Command14.Enabled = False
   If DERACCE%("ACTUPREC", "Actulización de Precios por C.Calculado") < 2 Then GoTo 99

   'SELECCION DE LISTA DE PRECIOS
   Call SELECHO("LISTAS/Selección de Lista para Actualizar")
   LPREX% = XVALO(VALACT1$("LISTAS"))
   If LPREX% < 1 Then GoTo 99
   '
   Timer1.Enabled = True
   '
99 Command14.Enabled = True
End Sub

Private Sub Command2_Click()
    '
    '09/04/07 (OT-07-0137)
    Call GRACONTROL("ANACOSTO", "PLANICOS", Text1)
    Call GRACONTROL("ANACOSTO", "PLACOSUP", Text2)

    Call GRACONTROL("ANACOSTO", "LIPREPES", Text7)
    Call GRACONTROL("ANACOSTO", "LIPREDOL", Text4)
    '
    Call GRACONTROL("ANACOSTO", "GAGENFAB", Text36.TEXT)
    Call GRACONTROL("ANACOSTO", "GAGENMAT", Text41.TEXT)
    Call GRACONTROL("ANACOSTO", "PORUTILI", Text39.TEXT)
    Call GRACONTROL("ANACOSTO", "PORIMPUE", Text43.TEXT)
    Call GRACONTROL("ANACOSTO", "PORMINOR", Me.Text8.TEXT) 'PORCENTAJE MINORISTA

    '
    TTXX$ = REGLEIDO$("CATOPER", 1)
    Call REPLA(TTXX$, MKS$(XVALO(Text5)), 29, 4)
    Call GRAREG("CATOPER", TTXX$, 1)
    Call CIERRARCH("CATOPER")
    '
    Call CIERRARCH("*.*")
    Hide
    '09/04/07 (FIN)
    '
End Sub

Private Sub Command21_Click()
    Command21.Enabled = False
    '
    Call TRALOCAL("GASTOGEN", "")
    VTB% = VENTABU%("GASTOGEN")
    If VTB% < 0 Then GoTo 99
    For U& = 1 To ULTREG&("!GASTOGEN")
      XX$ = REGLEIDO$("!GASTOGEN", U&)
      COD$ = TRIM$(Left$(XX$, 2))
      If COD$ = "" Then
         Call COMUNI("Falta Código en Fila: " & U&)
         GoTo 99
      End If
      
      PORCENTAJE% = CVI(Mid$(XX$, 3, 2))
      If PORCENTAJE% < 1 Then
         Call COMUNI("Falta Porcentaje en Fila: " & U&)
         GoTo 99
      End If
      
      DESDEX = CVS(Mid$(XX$, 5, 4))
      If DESDEX < 1 And U& > 1 Then
         Call COMUNI("Falta Valor Desde en Fila: " & U&)
         GoTo 99
      End If
    Next U&
    '
    Call TRACENTRAL("GASTOGEN", "")
    Call SETULTREG("!GASTOGEN", U&)
    Call CIERRARCH("!GASTOGEN")
    Call CIERRARCH("GASTOGEN")

    
99  Command21.Enabled = True
End Sub

Sub Command6_Click()
   Command6.Enabled = False
   SETUPCOM.Command11_Click
   Command6.Enabled = True
End Sub


Private Sub Command3_Click()
    '
    '09/04/07 (OT-07-0137)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Label4.Caption = VALACT2$("INDIFRM")
    End If
    '09/04/07 (FIN)
    '
End Sub
'

Private Sub Command4_Click()
    '
    '09/04/07 (OT-07-0137)
    Call SELECHO("LISTAS")
    VDEVU$ = VALACT1$("LISTAS")
    If VDEVU$ <> "" Then
        Text4.TEXT = VDEVU$
    End If
    '09/04/07 (FIN)
    '
End Sub

Private Sub Command5_Click()
    '
    '09/04/07 (OT-07-0137)
    Command5.Enabled = False
    Call CONECRUN("FLFORMS/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
    Command5.Enabled = True
    '09/04/07 (FIN)
    '
End Sub


Private Sub Command7_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text2.TEXT = VDEVU$
        Label7.Caption = VALACT2$("INDIFRM")
    End If

End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CONECRUN("FLFORMS/" + TRIM$(Text2.TEXT), "Formularios de Impresion")
    Command8.Enabled = True

End Sub

Private Sub Command9_Click()
    '
    '09/04/07 (OT-07-0137)
    Call SELECHO("LISTAS")
    VDEVU$ = VALACT1$("LISTAS")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
    End If
    '09/04/07 (FIN)
    '
End Sub

Private Sub Form_Load()
    '
    '09/04/07 (OT-07-0137)
    Text1 = Control$("ANACOSTO", "PLANICOS")
    Label4.Caption = ECOARCH$("INDIFRM", Text1)
    '
    Text2 = Control$("ANACOSTO", "PLACOSUP")
    Label7.Caption = ECOARCH$("INDIFRM", Text2)
    '
    Text7 = Control$("ANACOSTO", "LIPREPES")
    Label5.Caption = ECOARCH$("LISTAS", Chr$(XVALO(Text7)))
    Text4 = Control$("ANACOSTO", "LIPREDOL")
    Label6.Caption = ECOARCH$("LISTAS", Chr$(XVALO(Text4)))
    '
    DCC1$ = TRIM$(Control$("ANACOSTO", "GAGENMAT"))
    Text41 = FORMATNUM$(XVALO(DCC1$), "F6.2")
    '
    DCC2$ = TRIM$(Control$("ANACOSTO", "GAGENFAB"))
    Text36 = FORMATNUM$(XVALO(DCC2$), "F6.2")
    '
    DCC3$ = TRIM$(Control$("ANACOSTO", "PORUTILI"))
    Text39 = FORMATNUM$(XVALO(DCC3$), "F6.2")
    '
    DCC4$ = TRIM$(Control$("ANACOSTO", "PORIMPUE"))
    Text43 = FORMATNUM$(XVALO(DCC4$), "F6.2")
    '
    DCC5$ = TRIM$(Control$("ANACOSTO", "PORMINOR"))
    Text8 = FORMATNUM$(XVALO(DCC5$), "F6.2")
    '
    Text5 = FORMATNUM$(VALHORA(0), "F9.4")
    '
    '09/04/07 (FIN)
    '
End Sub

Private Sub mnuprobra_Click()
    'LEO LOS CONTROLES
    CANTIOP% = XVALO(Control("ANACOSJM", "CANOPPMO"))
    AUSENTISMO% = XVALO(Control("ANACOSJM", "AUSENTIS"))
    STOCX% = XVALO(Control("ANACOSJM", "STOCKJYM"))
    'CALLREPLO
    Call REPLA(TX$, MKI$(CANTIOP%), 1, 2)
    Call REPLA(TX$, MKI$(AUSENTISMO%), 3, 2)
    Call REPLA(TX$, MKI$(STOCX%), 5, 2)
    VDEF$ = TX$
    '
    'MUESTRO EL OBJETO
10  OBX$ = OBJPLA$("PARAMPRO", VDEF$)
    If OBX$ = "" Then
        Exit Sub
    End If
    '
    'LEO
    CANTIO$ = SAFETEXT$(CVI(Left$(OBX$, 2)))
    AUSENTI$ = SAFETEXT$(CVI(Mid$(OBX$, 3, 2)))
    STOCZZ$ = SAFETEXT$(CVI(Mid$(OBX$, 5, 2)))
    
    'GRABO
    Call GRACONTROL("ANACOSJM", "CANOPPMO", CANTIO$)
    Call GRACONTROL("ANACOSJM", "AUSENTIS", AUSENTI$)
    Call GRACONTROL("ANACOSJM", "STOCKJYM", STOCZZ$)

End Sub

Private Sub Text1_DblClick()
    '
    Call Command3_Click
    '
End Sub

Private Sub Text2_DblClick()
    Call Command7_Click
End Sub

Private Sub TEXT3_Change()
   
   Label11 = rasocli$(XVALO(Text3))
   Label122 = ""
   CONDI$ = "Nume_Cli = " & XVALO(Text3)
   Text6 = ""
   Text6 = VALOBADA("Deudor12", "CateMobra", CONDI$, "NOMESS")

End Sub

Private Sub Text4_Change()
    '
    Label6.Caption = ECOARCH$("LISTAS", Chr$(XVALO(Text4)))
    '
End Sub

Private Sub Text6_Change()
    For U& = 1 To ULTREG&("GASTOGEN")
      XX$ = REGLEIDO$("GASTOGEN", U&)
      COD$ = TRIM$(Left$(XX$, 2))
      If COD$ = TRIM$(Text6) Then
         PORCENTAJE% = CVI(Mid$(XX$, 3, 2))
         Exit For
      End If
    Next U&
   Label122.Caption = PORCENTAJE%
End Sub

Private Sub Text7_Change()
    '
    Label5.Caption = ECOARCH$("LISTAS", Chr$(XVALO(Text7)))
    '
End Sub

Private Sub Text7_DblClick()
    '
    Call Command9_Click
    '
End Sub

Private Sub Timer1_Timer()
    Static U%, CIXI1%(), CARGOVECTOR_CI%, TIPOMONEDA%, VUELTA%, MAXIVECTOR%, OPX%
    '
    If LPREX% < 1 Then Timer1.Enabled = False
    '
    If CARGOVECTOR_CI% < 1 Then
        Screen.MousePointer = 11
        'CARGO EL VECTOR DE LOS CODIGOS DE LA LISTA
        SQLX$ = "SELECT * FROM PREVENTA WITH(NOLOCK) WHERE COD_LISTA = " & LPREX%
        SQLX$ = SQLX$ + " ORDER BY Cod_Pieza"
        Dim LISTATMP As ADODB.Recordset
        Set LISTATMP = READSET(SQLX$)
        '
        L% = 0
        With LISTATMP
           Do While Not .EOF
              L% = L% + 1
              ReDim Preserve CIXI1%(L%)
              CIXI1%(L%) = XVALO(!Cod_Pieza)
              .MoveNext
           Loop
        End With
        '
        LISTATMP.Close
        Set LISTATMP = Nothing
        '
        'VERIFICO EL TIPO DE MONEDA PARA VER QUE LABEL LEO
        TIPOMONEDA% = XVALO(VALOBADA("LISTAPRE", "MONEDA", "COD_LISTA = " & LPREX%, "NOMESS"))
        CARGOVECTOR_CI% = 1 ' ASIGNO UNO PARA TENER DE BANDERA Y QUE NO RECORRA LA LISTA
        MAXIVECTOR% = UBound(CIXI1%) 'ASIGNO EL VALOR MAXIMO DEL VECTOR
    End If
    '
    If U% < 1 Then U% = 1 'ASIGNO VALOR POR DEFECTO PARA LA PRIMERA VEZ
    If VUELTA% > 5 Then U% = U% + 1: VUELTA% = 0
    If U% <= MAXIVECTOR% Then 'VALIDO QUE SEA <= AL MAXIMO DEL VECTOR PARA SABER SI YA RECORRIO TODOS
        Screen.MousePointer = 11
        ANACOSI09.Text1 = CODEXT$(CIXI1%(U%)) 'PONGO EL CODIGO EN PANTALLA
        PRECIO_ACTUAL# = XVALO(ANACOSI09.Label12) ' LEO EL COSTO CALCULADO
        If TIPOMONEDA% > 1 Then PRECIO_ACTUAL# = XVALO(ANACOSI09.Label3) 'POR SI ES LISTA DOLAR
        '
        If PRECIO_ACTUAL# > 0 Then 'SI NO LEE UN VALOR VUELVE A PASAR HASTA 5 VECES
           PRECI_LISTA# = ROUND#(PRELISTA#(TRIM$(Str$(LPREX%)), CIXI1%(U%)), 4) 'LEE EL PRECIO DE LISTA
           'SI EL PRECIO DEL LISTA ES MENOR AL CALCULADO ACTUALIZA
           If PRECI_LISTA# < PRECIO_ACTUAL# Then
              If OPX% < 3 Then
                    OPX% = COMALTER%("Se Ha Encontrado Diferencia de Precios para el Codigo '" & CODEXT$(CIXI1%(U%)) & "'\ \Costo Calculado: " & _
                                TRIM$(FORMATNUM$(PRECIO_ACTUAL#, "F16.4")) & "\Precio de Lista: " & TRIM$(FORMATNUM$(PRECI_LISTA#, "F16.4")) & "\ \" & _
                                "el Cual Será Actualizado a: " & TRIM$(FORMATNUM$(PRECIO_ACTUAL#, "F16.2")) & " \ \" & _
                                "¿Desea Actualizar la Lista de Precios?\\No, Conservar\Si, Actualizar\Actualizar Todos\Cancelar Actualización")
              End If
              If OPX% = 4 Then Timer1.Enabled = False: Call COMUNI("Proceso Interrumpido"): Command14.Enabled = True: OPX% = 0: Exit Sub
              If OPX% >= 2 Then
                CONDI$ = "COD_LISTA = " & LPREX% & " AND COD_PIEZA = " & CIXI1%(U%)
                Call ACTUREGISTRO("PREVENTA", "Precio_Unitario", CONDI$, FORMATNUM$(PRECIO_ACTUAL#, "F16.2"), REGAF&, "NOMESS")
              End If
           End If
           '
           VUELTA% = 0: U% = U% + 1
        Else
           'EN CASO QUE NO LEE IMPORTE DEL LABEL ACUMULA PARA PASAR HASTA 5 VECES
           VUELTA% = VUELTA% + 1
        End If
    Else
        'SI YA RECORRIO TODO EL VECTOR DESHABILITA EL TIMER Y TERMINA.
        U% = 0: OPX% = 0 'INICIALIZA POR QUE SON VARIABLE STATIC
        Timer1.Enabled = False
        Screen.MousePointer = 1
        Command14.Enabled = True
        Call COMUNI("Proceso Finalizado")
    End If

End Sub
Sub SeleccionMultipleCentroCostos()

  CONDI_PUBLIC$ = ""
  Dim FILAXX&()
  CANTICUENTAS% = XVALO(CantRegistros("PLANCUEN", , "NOMESS"))
  ReDim Preserve FILAXX&(CANTICUENTAS%, 3)  '
  '
  'GENERA LA CONDICION SEGUN LOS CENTRO DE COSTOS SELECCIONADOS PARA LA CONSULTA SQL
  With MUESTRADATOS1
     ATXX% = 0
     CONDI$ = ""
     For KKU& = 1 To .GRID.Rows - 1
       If .GRID.TextMatrix(KKU&, 0) = "*" Then
         CCOSTX% = XVALO(.GRID.TextMatrix(KKU&, 1))
         If ATXX% < 1 Then
              CONDI$ = CONDI$ + " (c.CENCOSTO = " & CCOSTX%
              ATXX% = 1
            Else
              CONDI$ = CONDI$ + " OR c.CENCOSTO = " & CCOSTX%
          End If
       End If
    Next KKU&
   End With
   '
   CONDI$ = CONDI$ + ") "
   If ATXX% < 1 Then Exit Sub

   'VALIDA SI SE SELECCIONARON C.COSTOS
   '+++++++++++++++++++++++++++++++++++++++++++++
   If CONDI$ = "" Then
     Call COMUNI("Falta Seleccionar Centro de Costos")
     GoTo 99
   End If
   CONDI_PUBLIC$ = CONDI$ ' LUEGO LA VOY A TENER QUE USAR EN MUESTRADATOS2
   '
   '
   '********************************************
   'CARGO EL LIST CON LOS MESES
    Call CARSELMESA
    Call SELECMESANO.LICOPEN.Clear
    For n& = 1 To ULTREG&("SELMESA")
       SELECMESANO.LICOPEN.AddItem REGLEIDO$("SELMESA", n&)
    Next n&
    SELECMESANO.Show 1
    '
    '*******************************************
    MESEX$ = "": ATXX% = 0: jH& = 0
    '´RECORRO LOS MESES SELECCIONADOS - ARMO LA CONSULTA PARA SQLX$ Y CARGO EN UNA VARIABLE LOS MESES-ENCABEZADOS
    '****************************************************************************************************************
    For KKU& = SELECMESANO.LICOPEN.ListCount To 1 Step -1
      jH& = jH& + 1
      
      If SELECMESANO.LICOPEN.Selected(KKU& - 1) = True Then
        ' CARGO EL STRING DE LOS CAMPOS PARA LOS ENCABEZADO
        TEXTOMES$ = TRIM$(SELECMESANO.LICOPEN.List(KKU& - 1))
        POS1% = InStr(TRIM$(TEXTOMES$), " ")
        MES$ = Left$(TEXTOMES$, POS1% - 1)
        ANO$ = TRIM$(Right(TEXTOMES$, Len(TEXTOMES$) - Len(MES$)))
        MESEX$ = MESEX$ + TEXTOMES$
        MESEX$ = MESEX$ + "/F12.2"
        MESEX$ = MESEX$ + ";"
        '*******************************************************
        'GENERO CONDICION SQL
        MESFILTRO$ = "01/" & MESNUM$(TEXTOMES$) & "/" & ANO$
        ANO1% = Right(FECHATEXLAR$(CVINT(MESFILTRO$)), 4)
        If ATXX% < 1 Then
             CONDI$ = CONDI$ + " AND (MONTH(C.FECHEMISI) = " & MESNUM$(TEXTOMES$)
             CONDI$ = CONDI$ + " AND  YEAR(C.FECHEMISI) = " & ANO1%
             ATXX% = 1
           Else
             CONDI$ = CONDI$ + " OR MONTH(C.FECHEMISI) = " & MESNUM$(TEXTOMES$)
             CONDI$ = CONDI$ + " AND  YEAR(C.FECHEMISI) = " & ANO1%
         End If
      End If
   Next KKU&
   CONDI$ = CONDI$ + ") "
   If ATXX% < 1 Then Exit Sub ' SI NO ELIGE NINGUN MES SE VA
   
    'CARGO ENCABEZADO DE GRILLA DE FILTRO
    Call CARGA_ENCABEZADO(MUESTRADATOS2.GRID, "Cta.Cble./A13;Denomincación/A24;" & MESEX$ & "Total/F14.2;CODIEMPR/F0", MUESTRADATOS2, 1, 1)
    '
    MUESTRADATOS2.MSF_2.Rows = 3
    MUESTRADATOS2.MSF_2.FixedRows = 2
    Call CARGA_ENCABEZADO(MUESTRADATOS2.MSF_2, "Cta.Cble./A13;Denomincación/A24;" & MESEX$ & "Total/F14.2;CODIEMPR/F0", MUESTRADATOS2, 1, 1)
    
    
    Dim Total()
    Erase Total
    ReDim Total(jH&)
    MUESTRADATOS2.GRID.Rows = 2
    VUELTA% = 0
    SQLX$ = "SELECT P.CODEXCUE ,  P.Denocue, C.CenCosto, SUM(C.IDebeComp) AS FACSUMAIDEBE, SUM(C.IHabeComp) AS NCSUMAHABER,"
    SQLX$ = SQLX$ + " MONTH(C.FECHEMISI) as mes, YEAR(C.FECHEMISI) AS ANO ,C.CodiEmpr,C.cuecontai"
    SQLX$ = SQLX$ + " FROM CAJABANC C  WITH(NOLOCK) INNER JOIN PLANCUEN P WITH(NOLOCK)"
    SQLX$ = SQLX$ + " ON C.cuecontai = P.CodInCue"
    'CONDICION
    SQLX$ = SQLX$ + " Where " & CONDI$
    '
    SQLX$ = SQLX$ + " GROUP BY  C.CenCosto,YEAR(C.FECHEMISI), MONTH(C.FECHEMISI),C.cuecontai, P.CODEXCUE , P.Denocue,C.CodiEmpr"
    SQLX$ = SQLX$ + " ORDER BY YEAR(C.FECHEMISI), MONTH(C.FECHEMISI),C.cuecontai, P.CODEXCUE, P.Denocue, C.CodiEmpr  "
    Set RS = READSET(SQLX$)
    With RS
    i% = 1
    Do While Not .EOF
     ' FORMATEO LA FECHA PARA QUE ME QUEDE COMO EL ENCABEZADO JULIO 15
      MEX = XVALO(!MES)  ' 7
      NUMAÑO = XVALO(!ANO) ' 2015
      MESTEX$ = TRIM$(METEXTO$(FECHANUM("01/" & Right$("00" & MEX, 2) & "/" + Right$(SAFETEXT$(NUMAÑO), 2)))) ' JULIO 2015
      MESTEX$ = Mid$(MESTEX$, 1, Len(MESTEX$) - 4) & Right$(SAFETEXT$(NUMAÑO), 2) 'JULIO 15
      '
      SUBRAYAR% = 0
      If VUELTA% = 0 Then FILAXX&(XVALO(!CUECONTAI), XVALO(!CODIEMPR)) = i%
      If ((SAFETEXT$(!CodExCue) <> CODEXCUE_ANT) Or (XVALO(!CODIEMPR) <> CodiEmpr_ANT)) And (VUELTA% > 0) Then
           i% = i% + 1
           MUESTRADATOS2.GRID.Rows = i%
           MUESTRADATOS2.GRID.Rows = MUESTRADATOS2.GRID.Rows + 1
           'USO UN VECTOR PARA LLEVAR EL CONTROL DE LA FILA YA QUE CUANDO CAMBIABA DE MES LA MISMA CUENTA LO IMPRIMIA EN OTRA FILA
           If FILAXX&(XVALO(!CUECONTAI), XVALO(!CODIEMPR)) = 0 Then
             FILAXX&(XVALO(!CUECONTAI), XVALO(!CODIEMPR)) = i%
           End If
           
      End If
      '
      If XVALO(!CODIEMPR) > 0 Then SUBRAYAR% = 1 ' PARA SUBRAYAR CASO MULTIEMPRESA  A MODO DE DIFERENCIAR UNO Y OTRO
      NUCOLUMNA% = columna%(MESTEX$, MUESTRADATOS2.GRID)
      MUESTRADATOS2.GRID.Row = FILAXX&(XVALO(!CUECONTAI), XVALO(!CODIEMPR)): MUESTRADATOS2.GRID.COL = 1
      MUESTRADATOS2.GRID.TextMatrix(FILAXX&(XVALO(!CUECONTAI), XVALO(!CODIEMPR)), 1) = SAFETEXT$(!CodExCue): If SUBRAYAR% > 0 Then MUESTRADATOS2.GRID.CellFontUnderline = True: MUESTRADATOS2.GRID.CellForeColor = vbBlue
      MUESTRADATOS2.GRID.TextMatrix(FILAXX&(XVALO(!CUECONTAI), XVALO(!CODIEMPR)), 2) = SAFETEXT$(!DENOCUE)
        IMPO# = XVALO(!FACSUMAIDEBE) - XVALO(!NCSUMAHABER)
      MUESTRADATOS2.GRID.TextMatrix(FILAXX&(XVALO(!CUECONTAI), XVALO(!CODIEMPR)), NUCOLUMNA%) = TRIM$(FORMATNUM$(IMPO#, "F12.2"))
      Total(NUCOLUMNA%) = Total(NUCOLUMNA%) + IMPO#
      MUESTRADATOS2.GRID.TextMatrix(FILAXX&(XVALO(!CUECONTAI), XVALO(!CODIEMPR)), MUESTRADATOS2.GRID.Cols - 1) = XVALO(!CODIEMPR)
      '
      'CUANDO CAMBIA LA CUENTA CONTABLE PASA AL SIGUIENTE REGISTRO
      CODEXCUE_ANT = TRIM$(!CodExCue)
      CodiEmpr_ANT = XVALO(!CODIEMPR)
      VUELTA% = 1
      .MoveNext
    Loop
    End With

    'Call COLOREAR_GRILLASQL_VERTICAL(MUESTRADATOS2.GRID, &HE0E0E0, 0, 2)
    MUESTRADATOS2.Caption = "Informe de Gastos - Centro de Costos: Selección Múltiple"
    Screen.MousePointer = 1
    MUESTRADATOS2.GRID.FixedCols = 3 '
    '
    'ORDENO POR LA DENOMINACION DE LA CUENTA.
    MUESTRADATOS2.GRID.COL = 1
    MUESTRADATOS2.GRID.Sort = 5
    
    'IMPRIMIR TOTALES VERTICALES

    i% = MUESTRADATOS2.GRID.Rows
    'TOTALIZO AL FINAL POR COLUMNA
    MUESTRADATOS2.GRID.Rows = i% + 2
    MUESTRADATOS2.GRID.TextMatrix(i%, 2) = "TOTALES: "
'    For j% = 3 To MUESTRADATOS2.GRID.Cols - 1 '
'       If j% > MUESTRADATOS2.GRID.Rows - 1 Then Exit For
'       MUESTRADATOS2.GRID.TextMatrix(I%, j%) = FORMATNUM$(Total(j%), "F12.2") ' MENOS 2 POR EL QUE EL VECTOR TIENE EL NUMERO REAL DEL MES
'    Next j%
    '
    'TOTALIZA DE MANERA VERTICAL
    Dim TOTALCOLU#()
    Erase TOTALCOLU#
    For L% = 1 To MUESTRADATOS2.GRID.Rows - 1
        For R% = 3 To MUESTRADATOS2.GRID.Cols - 1
          If L% >= MUESTRADATOS2.GRID.Rows - 2 Then Exit For
          MUESTRADATOS2.GRID.TextMatrix(MUESTRADATOS2.GRID.Rows - 2, R%) = FORMATNUM$(XVALO(MUESTRADATOS2.GRID.TextMatrix(MUESTRADATOS2.GRID.Rows - 2, R%)) + XVALO(MUESTRADATOS2.GRID.TextMatrix(L%, R%)), "F12.2")
        Next R%
    Next L%

    'TOTALIZA DE MANERA HORIZONTAL
    For L% = 1 To MUESTRADATOS2.GRID.Rows - 1
        TOTALFILA# = 0
        For R% = 3 To MUESTRADATOS2.GRID.Cols - 1
          TOTALFILA# = TOTALFILA# + XVALO(MUESTRADATOS2.GRID.TextMatrix(L%, R%))
        Next R%
        MUESTRADATOS2.GRID.TextMatrix(L%, MUESTRADATOS2.GRID.Cols - 2) = FORMATNUM$(TOTALFILA#, "F12.2")
    Next L%
    
    Call COLOREAR_GRILLASQL(MUESTRADATOS2.GRID, &HE0E0E0)

    'COLOREO LA FILA DE LOS TOTALES
    Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MUESTRADATOS2.GRID, &HC00000, i%)
    Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MUESTRADATOS2.GRID, &H808080, i%)
    '
    'POSICIONARSE AL FINAL DE LA LISTA
     TPXI& = MUESTRADATOS2.GRID.Rows - 22: If TPXI& < 1 Then TPXI& = 1
     On Error Resume Next
     MUESTRADATOS2.GRID.TopRow = TPXI&
     On Error GoTo 0
     '
    MUESTRADATOS2.ANO = SAFETEXT$(ANOSELECCIONADO%)
    MUESTRADATOS2.CENTROCOSTO = "Selección Múltiple"
    MUESTRADATOS2.Show 1
'    GoTo 20
    '
    'FALTA VER EL TEMA DE LOS TOTALES
99
   Command13.Enabled = True

End Sub

