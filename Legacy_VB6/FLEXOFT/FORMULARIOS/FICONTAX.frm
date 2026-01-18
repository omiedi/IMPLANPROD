VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form FICONTAX 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda Telefónica FLEXOFT "
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "FICONTAX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9510
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00D0FFFF&
      Height          =   2535
      Left            =   2730
      ScaleHeight     =   2475
      ScaleWidth      =   4155
      TabIndex        =   68
      Top             =   3255
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
         TabIndex        =   75
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
         TabIndex        =   74
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
         TabIndex        =   73
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
         TabIndex        =   72
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
         TabIndex        =   71
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
         TabIndex        =   70
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
         TabIndex        =   69
         Top             =   0
         Width           =   4110
      End
   End
   Begin VB.CommandButton Command14 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8840
      Picture         =   "FICONTAX.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Configuración de Funcionamiento"
      Top             =   4620
      Width           =   600
   End
   Begin VB.CommandButton Command13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8840
      Picture         =   "FICONTAX.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Emisión de Listados"
      Top             =   5280
      Width           =   600
   End
   Begin VB.CommandButton Command25 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8840
      Picture         =   "FICONTAX.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Imprime Ficha del Contacto"
      Top             =   5940
      Width           =   600
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
      Height          =   1270
      Left            =   75
      TabIndex        =   65
      Top             =   5280
      Width           =   8670
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   315
         Width           =   8520
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
      Height          =   1560
      Left            =   75
      TabIndex        =   39
      Top             =   2160
      Width           =   8670
      Begin VB.TextBox Text25 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         HelpContextID   =   240
         Left            =   6225
         TabIndex        =   13
         Text            =   " "
         Top             =   600
         Width           =   2310
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         HelpContextID   =   240
         Left            =   6225
         TabIndex        =   14
         Text            =   " "
         Top             =   885
         Width           =   1845
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         HelpContextID   =   240
         Left            =   1155
         TabIndex        =   10
         Text            =   " "
         Top             =   885
         Width           =   4320
      End
      Begin VB.TextBox Text16 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         HelpContextID   =   240
         Left            =   1155
         TabIndex        =   7
         Text            =   " "
         Top             =   315
         Width           =   4320
      End
      Begin VB.TextBox Text15 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         HelpContextID   =   240
         Left            =   1155
         TabIndex        =   8
         Top             =   600
         Width           =   1260
      End
      Begin VB.TextBox Text13 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         HelpContextID   =   240
         Left            =   3225
         TabIndex        =   9
         Text            =   " "
         Top             =   600
         Width           =   2250
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         HelpContextID   =   240
         Left            =   1155
         TabIndex        =   11
         Text            =   " "
         Top             =   1170
         Width           =   4320
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         HelpContextID   =   240
         Left            =   6225
         TabIndex        =   15
         Text            =   " "
         Top             =   1170
         Width           =   2310
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         HelpContextID   =   240
         Left            =   6225
         TabIndex        =   12
         Text            =   " "
         Top             =   315
         Width           =   1845
      End
      Begin VB.Image Image3 
         Height          =   300
         Left            =   8175
         Picture         =   "FICONTAX.frx":10C0
         Stretch         =   -1  'True
         Top             =   870
         Width           =   315
      End
      Begin VB.Image Image1 
         Height          =   300
         Left            =   8175
         Picture         =   "FICONTAX.frx":13CA
         Stretch         =   -1  'True
         Top             =   270
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
         Left            =   5250
         TabIndex        =   55
         Top             =   900
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
         Top             =   960
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
         Left            =   5250
         TabIndex        =   53
         Top             =   1200
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
         Left            =   150
         TabIndex        =   52
         Top             =   660
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
         Left            =   5250
         TabIndex        =   51
         Top             =   360
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
         Top             =   1200
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
         Left            =   2250
         TabIndex        =   49
         Top             =   660
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
         Top             =   360
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
      Height          =   2085
      Left            =   75
      TabIndex        =   38
      Top             =   75
      Width           =   8655
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
         Height          =   780
         Left            =   7350
         Picture         =   "FICONTAX.frx":16D4
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1200
         Width           =   555
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   7950
         Picture         =   "FICONTAX.frx":19DE
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   1200
         Width           =   555
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
         Height          =   620
         Left            =   7350
         TabIndex        =   66
         Top             =   225
         Width           =   1155
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   150
            TabIndex        =   67
            Top             =   280
            Width           =   840
         End
      End
      Begin VB.TextBox Text6 
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
         Left            =   1125
         TabIndex        =   5
         Text            =   " "
         Top             =   1462
         Width           =   6135
      End
      Begin VB.TextBox Text7 
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
         Left            =   1125
         TabIndex        =   6
         Text            =   " "
         Top             =   1725
         Width           =   6150
      End
      Begin VB.TextBox Text5 
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
         Left            =   1125
         TabIndex        =   4
         Text            =   " "
         Top             =   1200
         Width           =   6150
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1125
         TabIndex        =   3
         Text            =   " "
         Top             =   600
         Width           =   6120
      End
      Begin VB.TextBox Text2 
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
         Left            =   3570
         TabIndex        =   2
         Text            =   " "
         Top             =   315
         Width           =   3675
      End
      Begin VB.TextBox Text1 
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
         Left            =   1125
         TabIndex        =   1
         Text            =   " "
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "N. y Apellido:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1650
         TabIndex        =   47
         Top             =   375
         Width           =   1800
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Dir-elec:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   75
         TabIndex        =   46
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Mostrar:"
         BeginProperty Font 
            Name            =   "Arial"
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
         Top             =   900
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Empresa:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   75
         TabIndex        =   42
         Top             =   600
         Width           =   960
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Apel:"
         BeginProperty Font 
            Name            =   "Arial"
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
         Top             =   375
         Width           =   960
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
         Height          =   240
         Left            =   1125
         TabIndex        =   40
         Top             =   900
         Width           =   7380
      End
   End
   Begin VB.CommandButton Command21 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8840
      Picture         =   "FICONTAX.frx":22E0
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Alta de Nuevo Contacto Telefónico"
      Top             =   900
      Width           =   600
   End
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
      Height          =   600
      Left            =   8840
      Picture         =   "FICONTAX.frx":242A
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   225
      Width           =   600
   End
   Begin VB.CommandButton Command9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8840
      Picture         =   "FICONTAX.frx":2574
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Ordenamiento Alfabético"
      Top             =   2220
      Width           =   600
   End
   Begin VB.Frame Frame6 
      Caption         =   "Ex/I"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   8820
      TabIndex        =   76
      Top             =   2880
      Width           =   640
      Begin VB.CommandButton Command12 
         Caption         =   "Imp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   90
         TabIndex        =   33
         ToolTipText     =   "Importa Libreta de Direcciones (de C:\FLEXOFT)"
         Top             =   900
         Width           =   460
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Exp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   90
         TabIndex        =   34
         ToolTipText     =   "Exporta Direcciones para Outlook Express"
         Top             =   1200
         Width           =   460
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Res"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   90
         TabIndex        =   32
         ToolTipText     =   "Restaura Copia de Seguridad (Zip)"
         Top             =   540
         Width           =   460
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Bak"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   90
         TabIndex        =   31
         ToolTipText     =   "Confecciona Copia de Seguridad (Zip)"
         Top             =   240
         Width           =   460
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
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8840
      Picture         =   "FICONTAX.frx":287E
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Guarda los Cambios Realizados"
      Top             =   1575
      Width           =   600
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
      Height          =   1560
      Left            =   75
      TabIndex        =   56
      Top             =   3720
      Width           =   8670
      Begin VB.TextBox Text26 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6225
         TabIndex        =   21
         Text            =   " "
         Top             =   315
         Width           =   1845
      End
      Begin VB.TextBox Text24 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6225
         TabIndex        =   23
         Text            =   " "
         Top             =   885
         Width           =   1845
      End
      Begin VB.TextBox Text23 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1155
         TabIndex        =   20
         Text            =   " "
         Top             =   1170
         Width           =   4320
      End
      Begin VB.TextBox Text22 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3225
         TabIndex        =   18
         Text            =   " "
         Top             =   600
         Width           =   2250
      End
      Begin VB.TextBox Text21 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1155
         TabIndex        =   17
         Top             =   600
         Width           =   1260
      End
      Begin VB.TextBox Text20 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1155
         TabIndex        =   16
         Text            =   " "
         Top             =   315
         Width           =   4320
      End
      Begin VB.TextBox Text19 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1155
         TabIndex        =   19
         Text            =   " "
         Top             =   885
         Width           =   4320
      End
      Begin VB.TextBox Text18 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6225
         TabIndex        =   24
         Text            =   " "
         Top             =   1170
         Width           =   2310
      End
      Begin VB.TextBox Text17 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6225
         TabIndex        =   22
         Text            =   " "
         Top             =   600
         Width           =   2310
      End
      Begin VB.Image Image5 
         Height          =   300
         Left            =   8175
         Picture         =   "FICONTAX.frx":2B88
         Stretch         =   -1  'True
         Top             =   870
         Width           =   315
      End
      Begin VB.Image Image4 
         Height          =   300
         Left            =   8175
         Picture         =   "FICONTAX.frx":2E92
         Stretch         =   -1  'True
         Top             =   270
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
         Top             =   360
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
         Left            =   2250
         TabIndex        =   63
         Top             =   660
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
         Top             =   1200
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
         Left            =   5250
         TabIndex        =   61
         Top             =   360
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
         Top             =   600
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
         Left            =   5250
         TabIndex        =   59
         Top             =   900
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
         Top             =   900
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
         Left            =   5250
         TabIndex        =   57
         Top             =   1200
         Width           =   960
      End
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   12000
      Y1              =   8200
      Y2              =   8200
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   10690
      Picture         =   "FICONTAX.frx":319C
      Stretch         =   -1  'True
      Top             =   7640
      Width           =   4725
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
Attribute VB_Name = "FICONTAX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Call AUTOEXEC
   Call CENTRAFORM(Me)
    
   ' Al establecer 0 en InputLen se indica a MSComm que
   ' lea todo el contenido del búfer de entrada cuando
   ' se use la propiedad Input.
   MSComm1.InputLen = 0
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
   '
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
   Width = Frame1.Left + Frame1.Width
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
   Call CIERRARCH("*.*")
   Call FINAL("")
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



