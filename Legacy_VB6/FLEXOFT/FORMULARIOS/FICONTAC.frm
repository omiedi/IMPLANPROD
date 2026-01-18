VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form FICONTAC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda Telefónica FLEXOFT "
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "FICONTAC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
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
      Left            =   10935
      Picture         =   "FICONTAC.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Alta de Nuevo Contacto Telefónico"
      Top             =   1000
      Width           =   750
   End
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
      Left            =   10935
      Picture         =   "FICONTAC.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   210
      Width           =   750
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   10815
      Sorted          =   -1  'True
      TabIndex        =   78
      Top             =   735
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Sort"
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
      Left            =   10935
      Picture         =   "FICONTAC.frx":059E
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Ordenamiento Alfabético"
      Top             =   2580
      Width           =   750
   End
   Begin VB.Frame Frame6 
      Caption         =   "Exp/Imp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   10815
      TabIndex        =   77
      Top             =   3465
      Width           =   990
      Begin VB.CommandButton Command12 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   105
         TabIndex        =   33
         ToolTipText     =   "Importa Libreta de Direcciones (de C:\FLEXOFT)"
         Top             =   650
         Width           =   330
      End
      Begin VB.CommandButton Command10 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   525
         TabIndex        =   34
         ToolTipText     =   "Exporta Direcciones para Outlook Express"
         Top             =   650
         Width           =   330
      End
      Begin VB.CommandButton Command8 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   525
         TabIndex        =   32
         ToolTipText     =   "Restaura Copia de Seguridad (Zip)"
         Top             =   315
         Width           =   330
      End
      Begin VB.CommandButton Command7 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   105
         TabIndex        =   31
         ToolTipText     =   "Confecciona Copia de Seguridad (Zip)"
         Top             =   315
         Width           =   330
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00D0FFFF&
      Height          =   2535
      Left            =   4095
      ScaleHeight     =   2475
      ScaleWidth      =   4155
      TabIndex        =   69
      Top             =   3150
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton Command6 
         Caption         =   "X"
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
         Left            =   3885
         TabIndex        =   76
         Top             =   -10
         Width           =   270
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   1260
         TabIndex        =   75
         Top             =   2100
         Width           =   1485
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Levante el Receptor !!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   330
         Left            =   420
         TabIndex        =   74
         Top             =   1680
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFFFF&
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
         Left            =   210
         TabIndex        =   73
         Top             =   1365
         Width           =   3795
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "45241284"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   420
         TabIndex        =   72
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Llamando a:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   315
         TabIndex        =   71
         Top             =   420
         Width           =   1800
      End
      Begin VB.Label Label25 
         BackColor       =   &H00800000&
         Caption         =   "Monitor de Estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   0
         TabIndex        =   70
         Top             =   0
         Width           =   4110
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DTREnable       =   -1  'True
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H80000000&
      Caption         =   "Varios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   10815
      TabIndex        =   68
      Top             =   4620
      Width           =   990
      Begin VB.CommandButton Command25 
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
         Picture         =   "FICONTAC.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Imprime Ficha del Contacto"
         Top             =   1995
         Width           =   750
      End
      Begin VB.CommandButton Command13 
         Caption         =   "List"
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
         Picture         =   "FICONTAC.frx":0F12
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Emisión de Listados"
         Top             =   1155
         Width           =   750
      End
      Begin VB.CommandButton Command14 
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
         Picture         =   "FICONTAC.frx":121C
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   315
         Width           =   750
      End
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
      Left            =   10935
      Picture         =   "FICONTAC.frx":165E
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Guarda los Cambios Realizados"
      Top             =   1790
      Width           =   750
   End
   Begin VB.Frame Frame4 
      Caption         =   "Comentarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1700
      Left            =   105
      TabIndex        =   65
      Top             =   6510
      Width           =   10620
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1370
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   315
         Width           =   10620
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Laboral / Comercial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1800
      Left            =   105
      TabIndex        =   56
      Top             =   4620
      Width           =   10620
      Begin VB.TextBox Text26 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   7875
         TabIndex        =   21
         Text            =   " "
         Top             =   315
         Width           =   2150
      End
      Begin VB.TextBox Text24 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   7875
         TabIndex        =   23
         Text            =   " "
         Top             =   945
         Width           =   2150
      End
      Begin VB.TextBox Text23 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1155
         TabIndex        =   20
         Text            =   " "
         Top             =   1365
         Width           =   5895
      End
      Begin VB.TextBox Text22 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   3675
         TabIndex        =   18
         Text            =   " "
         Top             =   630
         Width           =   3375
      End
      Begin VB.TextBox Text21 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1155
         TabIndex        =   17
         Top             =   630
         Width           =   1485
      End
      Begin VB.TextBox Text20 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1155
         TabIndex        =   16
         Text            =   " "
         Top             =   315
         Width           =   5895
      End
      Begin VB.TextBox Text19 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1155
         TabIndex        =   19
         Text            =   " "
         Top             =   1050
         Width           =   5895
      End
      Begin VB.TextBox Text18 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   7875
         TabIndex        =   24
         Text            =   " "
         Top             =   1365
         Width           =   2535
      End
      Begin VB.TextBox Text17 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   7875
         TabIndex        =   22
         Text            =   " "
         Top             =   630
         Width           =   2535
      End
      Begin VB.Image Image5 
         Height          =   300
         Left            =   10080
         Picture         =   "FICONTAC.frx":1968
         Stretch         =   -1  'True
         Top             =   945
         Width           =   315
      End
      Begin VB.Image Image4 
         Height          =   300
         Left            =   10080
         Picture         =   "FICONTAC.frx":1C72
         Stretch         =   -1  'True
         Top             =   315
         Width           =   315
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Calle y Nº:"
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
         Left            =   105
         TabIndex        =   64
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Ciudad:"
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
         Left            =   2625
         TabIndex        =   63
         Top             =   735
         Width           =   960
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "País:"
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
         Left            =   105
         TabIndex        =   62
         Top             =   1470
         Width           =   960
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Telef:"
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
         Left            =   6825
         TabIndex        =   61
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "C.Post:"
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
         Left            =   105
         TabIndex        =   60
         Top             =   735
         Width           =   960
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Fax:"
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
         Left            =   6825
         TabIndex        =   59
         Top             =   1050
         Width           =   960
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Provincia:"
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
         Left            =   105
         TabIndex        =   58
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Cargo:"
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
         Left            =   6825
         TabIndex        =   57
         Top             =   1470
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Particular"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1800
      Left            =   105
      TabIndex        =   39
      Top             =   2730
      Width           =   10620
      Begin VB.TextBox Text25 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   7875
         TabIndex        =   13
         Top             =   630
         Width           =   2535
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   7875
         TabIndex        =   14
         Text            =   " "
         Top             =   1050
         Width           =   2150
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
         Height          =   280
         Left            =   1155
         TabIndex        =   10
         Text            =   " "
         Top             =   1050
         Width           =   5895
      End
      Begin VB.TextBox Text16 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1155
         TabIndex        =   7
         Text            =   " "
         Top             =   315
         Width           =   5895
      End
      Begin VB.TextBox Text15 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1155
         TabIndex        =   8
         Top             =   630
         Width           =   1485
      End
      Begin VB.TextBox Text13 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   3675
         TabIndex        =   9
         Text            =   " "
         Top             =   630
         Width           =   3375
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1155
         TabIndex        =   11
         Text            =   " "
         Top             =   1365
         Width           =   5895
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   7875
         TabIndex        =   15
         Text            =   " "
         Top             =   1365
         Width           =   2535
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   7875
         TabIndex        =   12
         Text            =   " "
         Top             =   315
         Width           =   2150
      End
      Begin VB.Image Image3 
         Height          =   300
         Left            =   10080
         Picture         =   "FICONTAC.frx":1F7C
         Stretch         =   -1  'True
         Top             =   1050
         Width           =   315
      End
      Begin VB.Image Image1 
         Height          =   300
         Left            =   10080
         Picture         =   "FICONTAC.frx":2286
         Stretch         =   -1  'True
         Top             =   315
         Width           =   315
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Celular:"
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
         Left            =   6825
         TabIndex        =   55
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Provincia:"
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
         Left            =   105
         TabIndex        =   54
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Fax:"
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
         Left            =   6825
         TabIndex        =   53
         Top             =   1470
         Width           =   960
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "C.Post:"
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
         Left            =   105
         TabIndex        =   52
         Top             =   735
         Width           =   960
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Telef:"
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
         Left            =   6825
         TabIndex        =   51
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "País:"
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
         Left            =   105
         TabIndex        =   50
         Top             =   1470
         Width           =   960
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Ciudad:"
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
         Left            =   2625
         TabIndex        =   49
         Top             =   735
         Width           =   960
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Calle y Nº:"
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
         Left            =   105
         TabIndex        =   48
         Top             =   420
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2535
      Left            =   105
      TabIndex        =   38
      Top             =   105
      Width           =   10620
      Begin VB.CommandButton Command11 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   460
         Left            =   8925
         Picture         =   "FICONTAC.frx":2590
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Búsqueda Secuencial por Texto"
         Top             =   1940
         Width           =   700
      End
      Begin VB.CommandButton Command5 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   460
         Left            =   8925
         Picture         =   "FICONTAC.frx":289A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Acceso Directo a Outlook Express"
         Top             =   1470
         Width           =   700
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   9720
         Picture         =   "FICONTAC.frx":2BA4
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   1470
         Width           =   700
      End
      Begin VB.Frame Frame5 
         Caption         =   "Nro"
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
         Left            =   8925
         TabIndex        =   66
         Top             =   210
         Width           =   1485
         Begin VB.Label Label26 
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
            Height          =   285
            Left            =   210
            TabIndex        =   67
            Top             =   315
            Width           =   1065
         End
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
         Height          =   280
         Left            =   1155
         TabIndex        =   5
         Text            =   " "
         Top             =   1785
         Width           =   7575
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
         Height          =   280
         Left            =   1155
         TabIndex        =   6
         Text            =   " "
         Top             =   2100
         Width           =   7575
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
         Height          =   280
         Left            =   1155
         TabIndex        =   4
         Text            =   " "
         Top             =   1470
         Width           =   7575
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
         Height          =   280
         Left            =   1155
         TabIndex        =   3
         Text            =   " "
         Top             =   630
         Width           =   7575
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
         Height          =   280
         Left            =   4095
         TabIndex        =   2
         Text            =   " "
         Top             =   315
         Width           =   4635
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
         Height          =   280
         Left            =   1155
         TabIndex        =   1
         Text            =   " "
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombre y Apellido:"
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
         Left            =   2205
         TabIndex        =   47
         Top             =   420
         Width           =   1800
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Dir-elec:"
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
         Left            =   105
         TabIndex        =   46
         Top             =   1575
         Width           =   960
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Mostrar:"
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
         Left            =   105
         TabIndex        =   45
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Empresa:"
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
         Left            =   105
         TabIndex        =   42
         Top             =   735
         Width           =   960
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Apel:"
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
         Left            =   105
         TabIndex        =   41
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label1 
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
         Height          =   285
         Left            =   1155
         TabIndex        =   40
         Top             =   1050
         Width           =   9255
      End
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   12000
      Y1              =   8200
      Y2              =   8200
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10350
      Picture         =   "FICONTAC.frx":34A6
      Top             =   7690
      Width           =   2010
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Apel:"
      Height          =   225
      Left            =   210
      TabIndex        =   44
      Top             =   1785
      Width           =   960
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Apel:"
      Height          =   225
      Left            =   210
      TabIndex        =   43
      Top             =   1260
      Width           =   960
   End
End
Attribute VB_Name = "FICONTAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command11_Click()
   '
   Static ARGUME$
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
10 ARGUME$ = TRIM$(UCase$(InputBox$("Buscar", "Búsqueda Cliente por Texto", ARGUME$)))
   '
   If ARGUME$ <> "" Then
     Screen.MousePointer = 11
     JJ& = 0
     For u& = 1 To ULTREG&("BCONTACT")
       XXX$ = REGLEIDO$("BCONTACT", u&)
       If InStr(UCase$(Mid$(XXX$, 3)), ARGUME$) > 0 Then
         JJ& = JJ& + 1
         Call GRAREG("!SECONTAC", XXX$, JJ&)
       End If
     Next u&
     '
     Call SETULTREG("!SECONTAC", JJ&)
     Call CIERRARCH("!SECONTAC")
     Call BAJALDISCO
     Call FRESHECHO("!SECONTAC")
     '
20   If JJ& = 1 Then
          TTXX1$ = REGLEIDO$("!SECONTAC", 1)
          NPX& = CVI(Left$(TTXX1$, 2))
          Label26 = Str$(NPX&)
        ElseIf JJ& > 1 Then
          Call SELECHO("!SECONTAC")
          NPX& = Val(VALACT1$("!SECONTAC"))
          If NPX& > 0 Then
            Label26 = Str$(NPX&)
          End If
        Else
          Call MENSERR(24, "Texto no Encontrado")
     End If
     '
   End If
   Screen.MousePointer = 1
   Command11.Enabled = True
End Sub

Private Sub Form_Load()
   Call AUTOEXEC
   Call CENTRAFORM(Me)
  
   ' Al establecer 0 en InputLen se indica a MSComm que
   ' lea todo el contenido del búfer de entrada cuando
   ' se use la propiedad Input.
   MSComm1.InputLen = 0
   If Screen.Width < 12000 Then
      Hide
      FICONTAX.Show
   End If
   '
End Sub

Private Sub Command1_Click()
   '
   Command1.Enabled = False
   If Val(Label26) > 0 Then
     Call GRAFICON(Name)
   End If
   Command1.Enabled = True
   '
End Sub

Private Sub Command10_Click()
   '
   Command10.Enabled = False
   Call EXPORAGE
   Command10.Enabled = True
End Sub

Private Sub Command12_Click()
   '
   Command12.Enabled = False
   Call IMPORAGE
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   If EXISTE%(LUBAS$ + "WINTABU.EXE") < 1 Then
     MsgBox "No Disponible en la Presente Versión."
   End If
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   SETUPTEL.Show 1
   Command14.Enabled = True
End Sub

Private Sub Command2_Click()
   If TRIM$(FORSELEC.Text1) = "" Then
     FORSELEC.ABUSQUE = TRIM$(Label26)
   End If
   Call SELECHO("BCONTACT")
   VDEVU$ = VALACT1$("BCONTACT")
   If VDEVU$ <> "" Then
     Label26 = TRIM$(VDEVU$)
   End If
End Sub

Private Sub COMMAND21_CLICK()
   Label26 = 1 + ULTREG&("BCONTACT")
   Text1.SetFocus
End Sub

Private Sub Command25_Click()
   ANCHAC = Width
   VISIAC = Picture1.Visible
   Image2.Visible = False
   Picture1.Visible = False
   Width = 10800
   On Error Resume Next
   PrintForm
   If Err Then
      On Error GoTo 0
      Call MENSERR(24, "Error de Impresión de Ficha")
   End If
   On Error GoTo 0
   Image2.Visible = True
   Picture1.Visible = VISIAC
   Width = ANCHAC
End Sub

Private Sub Command3_Click()
   FINAGETEL.Show 1
   If FINAGETEL.Label7 <> "REANUDA" Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     Call GENBACKUP
     Call FINAL("")
   End If
   'Call CIERRARCH("*.*")
   'Call FINAL("")
End Sub

Private Sub Command40_Click()
   COMMAND40.Enabled = False
   Call IMPOROUT
   COMMAND40.Enabled = True
End Sub

Private Sub Command4_Click()
   CancelFlag = 1
End Sub

Private Sub Command5_Click()
   '
   Command5.Enabled = False
   PASEL$ = CONTROL("AGETELEF", "RUT-OUT")
   FISEL$ = CONTROL("AGETELEF", "EXE-OUT")
   '
   On Error GoTo 20
   Shell PASEL$ + "\" + FISEL$, 1
   On Error GoTo 0
   Command5.Enabled = True
   Exit Sub
   '
20 Resume 21
21 On Error GoTo 0
   Command5.Enabled = True
   Call MENSERR(24, "Error de Direccionamiento")
   '
End Sub

Private Sub Command6_Click()
   CancelFlag = 1
End Sub

Private Sub Command7_Click()
   '
   Command7.Enabled = False
   Call CIERRARCH("BCONTACT")
   Call CIERRARCH("NOTACONT")
   Call CIERRARCH("*.*")
   '
   If TRIM$(UCase$(CurDir)) <> "C:\AGETEL" Then
     TTXX$ = "Su Agenda Telefónica FLEXOFT está Instalada"
     TTXX$ = TTXX$ + "\en la Carpeta " + REPLACAR$(CurDir, "\", "/") + "."
     TTXX$ = TTXX$ + "\  \Las Opciones de Respaldo y Recuperación"
     TTXX$ = TTXX$ + "\Pueden no Funcionar Correctamente en esta carpeta."
     TTXX$ = TTXX$ + "\  \Conviene Re-Instalar (Copiar) Toda la Agenda"
     TTXX$ = TTXX$ + "\a la Carpeta 'C:/AGETEL'."
     Call MENSERR(24, TTXX$)
   End If
   '
   If COMALTER%("Coloque Diskette de Respaldo\\Cancelar\Continuar") <> 2 Then
     GoTo 99
   End If
   '
   Screen.MousePointer = 11
   LDT$ = TRIM$(UCase$(LUDAT$))
   NX% = FILEOPEN%(LUCOM$ + "BKAGENTE.BAT", 2, 0)
   Print #NX%, LUBAS$ + "PKZIP " + "A:BCONTACT " + LDT$ + "BCONTACT.DAT " + LDT$ + "BCONTACT.X01 " + LDT$ + "NOTACONT.DAT " + LDT$ + "NOTACONT.X01 "
   Print #NX%, "EXIT"
   Close #NX%: NX% = 0
   Call BAJALDISCO
   '
   Shell LUCOM$ + "BKAGENTE.PIF", 2
   TINI = Timer
   While Timer < TINI + 5: Wend
   '
99 Command7.Enabled = True
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command8_Click()
   '
   Command8.Enabled = False
   Call CIERRARCH("BCONTACT")
   Call CIERRARCH("*.*")
   '
   If TRIM$(UCase$(CurDir)) <> "C:\AGETEL" Then
     TTXX$ = "Su Agenda Telefónica FLEXOFT está Instalada"
     TTXX$ = TTXX$ + "\en la Carpeta " + REPLACAR$(CurDir, "\", "/") + "."
     TTXX$ = TTXX$ + "\  \Las Opciones de Respaldo y Recuperación"
     TTXX$ = TTXX$ + "\Pueden no Funcionar Correctamente en esta carpeta."
     TTXX$ = TTXX$ + "\  \Conviene Re-Instalar (Copiar) Toda la Agenda"
     TTXX$ = TTXX$ + "\a la Carpeta 'C:/AGETEL'."
     Call MENSERR(24, TTXX$)
   End If
   '
   If COMALTER%("Coloque Diskette de Respaldo\\Cancelar\Continuar") <> 2 Then
     GoTo 99
   End If
   '
   If EXISTE%("A:BCONTACT.ZIP") <> 1 Then
     Call COMUNI("Recuperación Cancelada")
     GoTo 99
   End If
   '
   Screen.MousePointer = 11
   LDT$ = TRIM$(UCase$(LUDAT$))
   NX% = FILEOPEN%(LUCOM$ + "BKAGENTE.BAT", 2, 0)
   Print #NX%, LUBAS$ + "PKUNZIP -O " + "A:BCONTACT " + LDT$
   Print #NX%, "EXIT"
   Close #NX%: NX% = 0
   Call BAJALDISCO
   '
   Shell LUCOM$ + "BKAGENTE.PIF", 2
   TINI = Timer
   While Timer < TINI + 5: Wend
   '
99 Command8.Enabled = True
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command9_Click()
   '
   Command9.Enabled = False
   Call SORCONTAC
   Command9.Enabled = True
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 1
   FINAGETEL.Show 1
   If FINAGETEL.Label7 <> "REANUDA" Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     Call GENBACKUP
     Call FINAL("")
   End If
   'Call CIERRARCH("*.*")
   'Call FINAL("")
End Sub

Private Sub Image1_Click()
   Image1.Enabled = False
   Call Dial(TRIM$(Text10.Text), Name)
   Image1.Enabled = True
End Sub

Private Sub Image3_Click()
   Image3.Enabled = False
   Call Dial(TRIM$(Text14.Text), Name)
   Image3.Enabled = True
End Sub

Private Sub Image4_Click()
   Image4.Enabled = False
   Call Dial(TRIM$(Text26.Text), Name)
   Image4.Enabled = True
End Sub

Private Sub Image5_Click()
   Image5.Enabled = False
   Call Dial(TRIM$(Text24.Text), Name)
   Image5.Enabled = True
End Sub

Private Sub Label26_Change()
   '
   Call CARFICHA(Name)
   '
End Sub

Private Sub Text1_Change()
   Call ARMAMUESTRA(Name)
End Sub

Private Sub Text1_GotFocus()
   If Val(Label26) < 1 Then
     Command2.SetFocus
   End If
End Sub

Private Sub Text10_DblClick()
   Text10.Enabled = False
   Call Dial(TRIM$(Text10.Text), Name)
   Text10.Enabled = True
End Sub

Private Sub Text11_DblClick()
   Text11.Enabled = False
   Call Dial(TRIM$(Text11.Text), Name)
   Text11.Enabled = True
End Sub

Private Sub Text14_DblClick()
   Text14.Enabled = False
   Call Dial(TRIM$(Text14.Text), Name)
   Text14.Enabled = True
End Sub

Private Sub Text17_DblClick()
   Text17.Enabled = False
   Call Dial(TRIM$(Text17.Text), Name)
   Text17.Enabled = True
End Sub

Private Sub Text2_Change()
   Call ARMAMUESTRA(Name)
End Sub

Private Sub Text24_DblClick()
   Text24.Enabled = False
   Call Dial(TRIM$(Text24.Text), Name)
   Text24.Enabled = True
End Sub

Private Sub Text25_DblClick()
   Text25.Enabled = False
   Call Dial(TRIM$(Text25.Text), Name)
   Text25.Enabled = True
End Sub

Private Sub Text26_DblClick()
   Text26.Enabled = False
   Call Dial(TRIM$(Text26.Text), Name)
   Text26.Enabled = True
End Sub

Private Sub Text4_Change()
   Call ARMAMUESTRA(Name)
End Sub


Private Sub Text5_DblClick()
   '
   NPX& = Val(Label26)
   If NPX& > 0 Then
      DIRELE$ = TRIM$(Text5)
      If DIRELE$ <> "" Then
         If InStr(DIRELE$, "@") > 1 Then
            Call CREAMAIL(NPX&)
            '
            ATENTI% = 1
            MAILMESS.Label26 = Label26
            MAILMESS.Show 1
         End If
      End If
   End If
   '
End Sub

'
'   REG& = REGBUS&("BCONTACT", 1, MKI$(IFICHA%))
'   If REG& < 1 Or REG& > ULTREG&("BCONTACT") Then
'       Forms(KI% - 1).Text1 = ""
'       Forms(KI% - 1).Text2 = ""
'       Forms(KI% - 1).Text4 = ""
'       Forms(KI% - 1).Text5 = ""
'       Forms(KI% - 1).Text6 = ""
'       Forms(KI% - 1).Text7 = ""
'       Forms(KI% - 1).Text8 = ""
'       Forms(KI% - 1).Text9 = ""
'       Forms(KI% - 1).Text16 = ""
'       Forms(KI% - 1).Text15 = ""
'       Forms(KI% - 1).Text13 = ""
'       Forms(KI% - 1).Text8 = ""
'       Forms(KI% - 1).Text12 = ""
'       Forms(KI% - 1).Text10 = ""
'       Forms(KI% - 1).Text25 = ""
'       Forms(KI% - 1).Text14 = ""
'       Forms(KI% - 1).Text11 = ""
'       Forms(KI% - 1).Text20 = ""
'       Forms(KI% - 1).Text21 = ""
'       Forms(KI% - 1).Text22 = ""
'       Forms(KI% - 1).Text19 = ""
'       Forms(KI% - 1).Text23 = ""
'       Forms(KI% - 1).Text26 = ""
'       Forms(KI% - 1).Text17 = ""
'       Forms(KI% - 1).Text24 = ""
'       Forms(KI% - 1).Text18 = ""
'     Else
'       RELE$ = REGLEIDO$("BCONTACT", REG&)
'       Forms(KI% - 1).Text1 = TRIM$(Mid$(RELE$, 81, 8))
'       Forms(KI% - 1).Text2 = TRIM$(Mid$(RELE$, 89, 56))
'       Forms(KI% - 1).Text4 = TRIM$(Mid$(RELE$, 145, 64))
'       Forms(KI% - 1).Text5 = TRIM$(Mid$(RELE$, 217, 64))
'       Forms(KI% - 1).Text6 = TRIM$(Mid$(RELE$, 281, 64))
'       Forms(KI% - 1).Text7 = TRIM$(Mid$(RELE$, 345, 64))
'       '
'       Forms(KI% - 1).Text16 = TRIM$(Mid$(RELE$, 512 + 1, 48))
'       Forms(KI% - 1).Text15 = TRIM$(Mid$(RELE$, 512 + 49, 16))
'       Forms(KI% - 1).Text13 = TRIM$(Mid$(RELE$, 512 + 65, 48))
'       Forms(KI% - 1).Text8 = TRIM$(Mid$(RELE$, 512 + 113, 32))
'       Forms(KI% - 1).Text12 = TRIM$(Mid$(RELE$, 512 + 145, 32))
'       Forms(KI% - 1).Text10 = TRIM$(Mid$(RELE$, 512 + 177, 20))
'       Forms(KI% - 1).Text25 = TRIM$(Mid$(RELE$, 512 + 197, 20))
'       Forms(KI% - 1).Text14 = TRIM$(Mid$(RELE$, 512 + 217, 20))
'       Forms(KI% - 1).Text11 = TRIM$(Mid$(RELE$, 512 + 237, 20))
'       '
'       Forms(KI% - 1).Text20 = TRIM$(Mid$(RELE$, 768 + 1, 48))
'       Forms(KI% - 1).Text21 = TRIM$(Mid$(RELE$, 768 + 49, 16))
'       Forms(KI% - 1).Text22 = TRIM$(Mid$(RELE$, 768 + 65, 48))
'       Forms(KI% - 1).Text19 = TRIM$(Mid$(RELE$, 768 + 113, 32))
'       Forms(KI% - 1).Text23 = TRIM$(Mid$(RELE$, 768 + 145, 32))
'       Forms(KI% - 1).Text26 = TRIM$(Mid$(RELE$, 768 + 177, 20))
'       Forms(KI% - 1).Text17 = TRIM$(Mid$(RELE$, 768 + 197, 20))
'       Forms(KI% - 1).Text24 = TRIM$(Mid$(RELE$, 768 + 217, 20))
'       Forms(KI% - 1).Text18 = TRIM$(Mid$(RELE$, 768 + 237, 20))
'       '
'
'
Sub CREAMAIL(NPX&)
   '
   IFICHA% = NPX&
   REG& = REGBUS&("BCONTACT", 1, MKI$(IFICHA%))
   If REG& < 1 Or REG& > ULTREG&("BCONTACT") Then Exit Sub
   '
   RELE$ = REGLEIDO$("BCONTACT", REG&)
   '
   EMPRE$ = TRIM$(Mid$(RELE$, 145, 64))
   CALNUM$ = TRIM$(Mid$(RELE$, 512 + 1, 48))
      If CALNUM$ = "" Then CALNUM$ = Text20
   CPOST$ = TRIM$(Mid$(RELE$, 512 + 49, 16))
      If CPOST$ = "" Then CPOST$ = Text21
   If CPOST$ <> "" Then CPOST$ = CPOST$ + " "
   LOCAX$ = TRIM$(Mid$(RELE$, 512 + 65, 48))
      If LOCAX$ = "" Then LOCAX$ = Text22
   LOCAX$ = CPOST$ + LOCAX$
   PROPA$ = TRIM$(Mid$(RELE$, 512 + 113, 32))
      If PROPA$ = "" Then PROPA$ = Text19
   If PROPA$ <> "" Then PROPA$ = PROPA$ + " - "
   PROPA$ = TRIM$(PROPA$ + TRIM$(Mid$(RELE$, 12 + 145, 32)))
      If TRIM$(Mid$(RELE$, 12 + 145, 32)) = "" Then PROPA$ = TRIM$(PROPA$ + TRIM$(Text23))
   '
   CADEST% = 0
   TTXX$ = TRIM$(Mid$(RELE$, 89, 56))
   If TTXX$ <> "" Then
      CADEST% = CADEST% + 1
      DESTIX$ = TTXX$
   End If
   '
   TTXX$ = ""
   If EMPRE$ <> "" Then TTXX$ = TTXX$ + EMPRE$ + Chr$(13) + Chr$(10)
   If CALNUM$ <> "" Then TTXX$ = TTXX$ + CALNUM$ + Chr$(13) + Chr$(10)
   If LOCAX$ <> "" Then TTXX$ = TTXX$ + LOCAX$ + Chr$(13) + Chr$(10)
   If PROPA$ <> "" Then TTXX$ = TTXX$ + PROPA$ + Chr$(13) + Chr$(10)
   '
   MMT4$ = TRIM$(DESTIX$)
   If MMT4$ <> "" And EMPRE$ <> "" Then MMT4$ = MMT4$ + " - "
   MAILMESS.Text4 = TRIM$(MMT4$ + EMPRE$)
     PPXX% = InStr(Text5, "/")
     If PPXX% < 1 Then PPXX% = 1 + Len(Text5)
   MAILMESS.Text5 = TRIM$(Left$(Text5, PPXX% - 1))
   MAILMESS.Text20 = "FLEXOFT"
   '
   HOII% = HOY(HOS$)
   MAILMESS.Label5 = "Buenos Aires, " + FETEXTO$(HOII%) + ".-"
   While MAILMESS.Label5.Width < MAILMESS.Text2.Width - 1500
     MAILMESS.Label5 = " " + MAILMESS.Label5
   Wend
   '
   TTYY$ = MAILMESS.Label5 + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10)
   If CADEST% < 1 Then
       TTYY$ = TTYY$ + "Sres." + Chr$(13) + Chr$(10)
    ElseIf TRIM$(Text1) <> "" Then
       TTYY$ = TTYY$ + TRIM$(Text1) + Chr$(13) + Chr$(10)
       TTYY$ = TTYY$ + DESTIX$ + Chr$(13) + Chr$(10)
    Else
       TTYY$ = TTYY$ + "Sr./Sra." + Chr$(13) + Chr$(10)
       TTYY$ = TTYY$ + DESTIX$ + Chr$(13) + Chr$(10)
   End If
   '
   TTYY$ = TTYY$ + TTXX$
   SST = 4 + Len(TTYY$)
   '
   For KKII% = 1 To 5: TTYY$ = TTYY$ + Chr$(13) + Chr$(10): Next KKII%
   TTYY$ = TTYY$ + LOADFILE$(LUDAT$ + "PIEDEMAI.DAT")
   '
   MAILMESS.Text2 = TTYY$ ' = TTYY$ + Text1
   MAILMESS.Text2.SelStart = SST
   MAILMESS.Text2.SelLength = 0
   '
End Sub


