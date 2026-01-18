VERSION 5.00
Begin VB.Form AJUSTOCK 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ajuste de Stocks a Fecha"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10155
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   10155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFF80&
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
      Height          =   1800
      Left            =   315
      TabIndex        =   40
      Top             =   2940
      Width           =   4110
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
         Height          =   285
         Left            =   1365
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   420
         Width           =   2640
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
         Left            =   1365
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   840
         Width           =   2640
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2625
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1260
         Width           =   1170
      End
      Begin VB.CommandButton Command5 
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
         TabIndex        =   41
         Top             =   1260
         Width           =   175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito"
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
         Left            =   315
         TabIndex        =   47
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   225
         Left            =   105
         TabIndex        =   46
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha del Ajuste:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   45
         Top             =   1260
         Width           =   2220
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Consultas - Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   4830
      TabIndex        =   23
      Top             =   1300
      Width           =   5160
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   210
         ScaleHeight     =   75
         ScaleWidth      =   4680
         TabIndex        =   49
         Top             =   4620
         Width           =   4740
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   50
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   210
         TabIndex        =   39
         Top             =   3465
         Width           =   3690
         Begin VB.OptionButton Option4 
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
            Height          =   225
            Left            =   2100
            TabIndex        =   35
            Top             =   315
            Width           =   1275
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Pantalla"
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
            Left            =   315
            TabIndex        =   34
            Top             =   315
            Value           =   -1  'True
            Width           =   1380
         End
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
         Height          =   600
         Left            =   4095
         Picture         =   "AJUSTOCK.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3600
         Width           =   855
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Rango Numérico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   210
         TabIndex        =   27
         Top             =   735
         Width           =   4740
         Begin VB.TextBox Text14 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   945
            TabIndex        =   28
            Top             =   420
            Width           =   1065
         End
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3255
            TabIndex        =   29
            Top             =   420
            Width           =   1065
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
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
            TabIndex        =   38
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
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
            Left            =   2415
            TabIndex        =   37
            Top             =   420
            Width           =   750
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Rango de Fechas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   210
         TabIndex        =   24
         Top             =   2100
         Width           =   4740
         Begin VB.TextBox Text12 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3255
            TabIndex        =   32
            Text            =   " "
            Top             =   420
            Width           =   1065
         End
         Begin VB.TextBox Text11 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   945
            TabIndex        =   30
            Text            =   " "
            Top             =   420
            Width           =   1065
         End
         Begin VB.CommandButton BOTSEL 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1995
            Picture         =   "AJUSTOCK.frx":030A
            TabIndex        =   31
            Top             =   420
            Width           =   175
         End
         Begin VB.CommandButton BOTSEL 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   4305
            Picture         =   "AJUSTOCK.frx":0614
            TabIndex        =   33
            Top             =   420
            Width           =   175
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
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
            Left            =   2415
            TabIndex        =   26
            Top             =   420
            Width           =   750
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
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
            TabIndex        =   25
            Top             =   420
            Width           =   855
         End
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Kardex"
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
      Picture         =   "AJUSTOCK.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2835
      TabIndex        =   1
      Top             =   210
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
      Height          =   645
      Left            =   315
      MultiLine       =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   630
      Width           =   4110
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
      Left            =   3780
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   210
      Width           =   645
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
      Left            =   1155
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   210
      Width           =   1695
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   315
      TabIndex        =   2
      Top             =   1470
      Width           =   4110
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
      Height          =   855
      Left            =   9135
      Picture         =   "AJUSTOCK.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command29 
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
      Left            =   8190
      Picture         =   "AJUSTOCK.frx":0D72
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command26 
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
      Height          =   855
      Left            =   4830
      Picture         =   "AJUSTOCK.frx":107C
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Direct"
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
      Picture         =   "AJUSTOCK.frx":14BE
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   210
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Ingreso de Ajustes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   315
      TabIndex        =   16
      Top             =   4935
      Width           =   4110
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
         Left            =   1785
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   945
         Width           =   1170
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
         Left            =   1785
         TabIndex        =   4
         Top             =   1365
         Width           =   1170
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
         Left            =   1785
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   525
         Width           =   1170
      End
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
         Height          =   285
         Left            =   1785
         TabIndex        =   5
         Top             =   1785
         Width           =   1170
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
         Height          =   705
         Left            =   3150
         Picture         =   "AJUSTOCK.frx":17C8
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   525
         Width           =   750
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
         Height          =   705
         Left            =   3150
         Picture         =   "AJUSTOCK.frx":1AD2
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1365
         Width           =   750
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Entrada:"
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
         TabIndex        =   21
         Top             =   945
         Width           =   1590
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Salida:"
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
         TabIndex        =   20
         Top             =   1365
         Width           =   1590
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "StockTeórico:"
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
         TabIndex        =   19
         Top             =   525
         Width           =   1590
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Ajustado:"
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
         TabIndex        =   18
         Top             =   1785
         Width           =   1590
      End
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "        Sistema Modular Integrado              Linea FLEXOFT (r) - Version 9.10"
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
      Left            =   6825
      TabIndex        =   48
      Top             =   6675
      Width           =   3315
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   4830
      Picture         =   "AJUSTOCK.frx":1DDC
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   5355
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "U-M"
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
      Left            =   2730
      TabIndex        =   13
      Top             =   210
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Height          =   225
      Left            =   105
      TabIndex        =   12
      Top             =   210
      Width           =   855
   End
End
Attribute VB_Name = "AJUSTOCK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTSEL_Click(Index As Integer)
    Call SELECHO("SELFECHA")
    If VALACT1$("SELFECHA") <> "" Then
        If Index = 0 Then
            Text11.TEXT = VALACT1$("SELFECHA")
          Else
            Text12.TEXT = VALACT1$("SELFECHA")
        End If
    End If
End Sub

Private Sub Command1_Click()
    PREPA% = 0
    AJUSTOCK.Hide
End Sub

Private Sub Command2_Click()
    COD$ = ""
    CI% = 0
    Text1.TEXT = ""
    Text2.TEXT = ""
    Text3.TEXT = ""
    Text6.TEXT = ""
    Text7.TEXT = ""
    Text8.TEXT = ""
    Text9.TEXT = ""
    Text1.SetFocus
End Sub

Private Sub Command26_Click()
    OPTRAJUST.Show 1
End Sub

Private Sub Command27_Click()
    Call ACCDIR("AJUSTOCK")
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("AJUSTOCK")
End Sub

Private Sub Command3_Click()
    Call AJUSTO3
    Call AJUSTO2
    Text6.TEXT = Text9.TEXT
    Text7.TEXT = ""
    Text8.TEXT = ""
    Text1.SetFocus
End Sub

Private Sub Command4_Click()
    '
    Call SELECHO("MASTER")
    COD$ = VALACT1$("MASTER")
    Text1.TEXT = COD$
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = UNIMED$(CI%)
    '
End Sub
'

Private Sub Command5_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text10.TEXT = VDEV$
    End If
End Sub

Private Sub Command6_Click()
    Command3.Enabled = False
    'Call CONECRUN("*KARDEX/" + COD$, "")
    Call GENKARDEX(COD$)
    Command3.Enabled = True
End Sub

Private Sub Command7_Click()
    '
    DESNUM& = 1: If Val(Text14.TEXT) > 0 Then DESNUM& = Val(Text14.TEXT)
    HASNUM& = 999999
    If TRIM$(Text13.TEXT) <> "" Then
        If Val(Text13.TEXT) >= DESNUM& Then HASNUM& = Val(Text13.TEXT)
    End If
    DESFE% = 0: If FECHANUM(TRIM$(Text11.TEXT)) > 0 Then DESFE% = FECHANUM(TRIM$(Text11.TEXT))
    HASFE% = 32767
    If TRIM$(Text12.TEXT) <> "" Then
        If FECHANUM(TRIM$(Text12.TEXT)) >= DESFE% Then HASFE% = FECHANUM(TRIM$(Text12.TEXT))
    End If
    '
    Screen.MousePointer = 11
    Call BLANARCH("LISTRAJ")
    FIN& = ULTREG&("AJUSTOCK")
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        X$ = REGLEIDO$("AJUSTOCK", I&)
        FF% = CVI(Left$(X$, 2))
        NRO& = CVS(Mid$(X$, 3, 4))
        If FF% >= DESFE% Then
          If FF% <= HASFE% Then
            NRO& = CVS(Mid$(X$, 3, 4))
            If NRO& >= DESNUM& Then
              If NRO& <= HASNUM& Then
                REFE$ = Mid$(X$, 34, 47)
                CI1% = CVI(Mid$(X$, 7, 2))
                DP1% = Asc(Mid$(X$, 33, 1))
                CAN# = CVD(Mid$(X$, 25, 8))
                Y$ = REGBLAN$("LISTRAJ")
                Call REPLA(Y$, MKI$(FF%), 1, 2)
                Call REPLA(Y$, REFE$, 3, 48)
                Call REPLA(Y$, MKS$(NRO&), 51, 4)
                Call REPLA(Y$, MKI$(CI1%), 55, 2)
                Call REPLA(Y$, UNIMED$(CI1%), 57, 4)
                If CAN# > 0 Then
                    Call REPLA(Y$, MKD$(CAN#), 61, 8)
                    Call REPLA(Y$, MKD$(0), 69, 8)
                  Else
                    Call REPLA(Y$, MKD$(0), 61, 8)
                    Call REPLA(Y$, MKD$(Abs(CAN#)), 69, 8)
                End If
                Call REPLA(Y$, Chr$(DP1%), 80, 1)
                Call REGAPP("LISTRAJ", Y$)
              End If
            End If
          End If
        End If
    Next I&
    Call CIERRARCH("LISTRAJ")
    Screen.MousePointer = 1
    '
    OPCI% = 1
    If Option4.Value = True Then OPCI% = 2
    If OPCI% = 1 Then
        Call OPNOIN("")
        'Call CONECRUN("*COTRAST", "")
      ElseIf OPCI% = 2 Then
        Call CONECRUN("*LIAJUST", "")
    End If
    '

End Sub

Private Sub Form_Load()
    Call AJUSTO1
End Sub
'

Private Sub List1_Click()
    '
    ILISTA& = List1.ListIndex
    Text4.TEXT = Left$(List1.TEXT, 20)
    Text6.TEXT = TRIM$(Mid$(List1.TEXT, 21))
    '
    AJUSTOCK.Text7.TEXT = ""
    AJUSTOCK.Text8.TEXT = ""
    AJUSTOCK.Text9.TEXT = TRIM$(Mid$(AJUSTOCK.List1.TEXT, 21))
    If Val(AJUSTOCK.Text9.TEXT) < 0 Then
        AJUSTOCK.Text7.TEXT = TRIM$(CSTRING$(MKD$((-1) * Val(AJUSTOCK.Text9.TEXT)), 4, 10, 1, 2))
        AJUSTOCK.Text9.TEXT = "0.0"
    End If
    Text9.SetFocus
End Sub

Private Sub Text1_Change()
    '
    Text2.TEXT = ""
    Text3.TEXT = ""
    If CODINT%(Text1.TEXT) > 0 Then
        COD$ = TRIM$(Text1.TEXT)
        CI% = CODINT%(COD$)
        Text2.TEXT = UNIMED$(CI%)
        Text3.TEXT = DESCRIT$(CI%)
        Call AJUSTO2               ' recalcular existencias a fecha
    End If
    '
End Sub
'

Private Sub Text1_DblClick()
    '
    Call SELECHO("MASTER")
    COD$ = VALACT1$("MASTER")
    Text1.TEXT = COD$
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = UNIMED$(CI%)
    '
End Sub
'

Private Sub Text1_LostFocus()
    COD$ = TRIM$(Text1.TEXT)
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = UNIMED$(CI%)
End Sub

Private Sub Text10_Change()
    If AJUSTOCK.Visible = True Then
        If FECHANUM(Text10.TEXT) > 0 Then
            Call AJUSTO2
        End If
    End If
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text10.TEXT = VDEV$
    End If
End Sub

Private Sub Text11_DblClick()
    Call SELECHO("SELFECHA")
    If VALACT1$("SELFECHA") <> "" Then
        Text11.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text12_DblClick()
    Call SELECHO("SELFECHA")
    If VALACT1$("SELFECHA") <> "" Then
        Text12.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text2_GotFocus()
    Text7.SetFocus

End Sub

Private Sub Text3_GotFocus()
    Text7.SetFocus

End Sub

Private Sub Text4_GotFocus()
    Text7.SetFocus

End Sub

Private Sub Text6_GotFocus()
    COD$ = Text1.TEXT
    If CODINT%(COD$) < 1 Then
        Text1.SetFocus
        Exit Sub
    End If
    ILISTA& = List1.ListIndex
    If ILISTA& < 0 Then
        List1.SetFocus
        Exit Sub
    End If
    Text9.SetFocus
End Sub
'

Private Sub Text7_Change()
    If MODENTRAJU% = 7 Then
        Text8.TEXT = ""
        CAJUS# = Val(Text6.TEXT) + Val(Text7.TEXT) - Val(Text8.TEXT)
        Text9.TEXT = Str$(CAJUS#)
    End If
End Sub
'

Private Sub Text8_Change()
    If MODENTRAJU% = 8 Then
        Text7.TEXT = ""
        CAJUS# = Val(Text6.TEXT) + Val(Text7.TEXT) - Val(Text8.TEXT)
        Text9.TEXT = Str$(CAJUS#)
    End If
End Sub
'

Private Sub Text9_Change()
    If MODENTRAJU% = 9 Then
       DIFER# = Val(Text9.TEXT) - Val(Text6.TEXT)
       Text7.TEXT = ""
       Text8.TEXT = ""
       If DIFER# > 0 Then
           Text7.TEXT = Str$(DIFER#)
         ElseIf DIFER# < 0 Then
           Text8.TEXT = Str$(Abs(DIFER#))
       End If
    End If
End Sub
'
Private Sub Text7_GotFocus()
    COD$ = Text1.TEXT
    If CODINT%(COD$) < 1 Then
        Text1.SetFocus
        Exit Sub
    End If
    ILISTA& = List1.ListIndex
    If ILISTA& < 0 Then
        List1.SetFocus
        Exit Sub
    End If
    Text7.SelStart = 0
    Text7.SelLength = Len(Text7.TEXT)
    MODENTRAJU% = 7
End Sub

Private Sub Text8_GotFocus()
    COD$ = Text1.TEXT
    If CODINT%(COD$) < 1 Then
        Text1.SetFocus
        Exit Sub
    End If
    ILISTA& = List1.ListIndex
    If ILISTA& < 0 Then
        List1.SetFocus
        Exit Sub
    End If
    Text8.SelStart = 0
    Text8.SelLength = Len(Text8.TEXT)
    MODENTRAJU% = 8
End Sub
'

Private Sub Text9_GotFocus()
    COD$ = Text1.TEXT
    If CODINT%(COD$) < 1 Then
        Text1.SetFocus
        Exit Sub
    End If
    ILISTA& = List1.ListIndex
    If ILISTA& < 0 Then
        List1.SetFocus
        Exit Sub
    End If
    Text9.SelStart = 0
    Text9.SelLength = Len(Text9.TEXT)
    MODENTRAJU% = 9
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command3.SetFocus
    End If
End Sub
